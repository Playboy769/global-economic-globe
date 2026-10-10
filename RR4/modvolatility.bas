Attribute VB_Name = "modvolatility"
Option Explicit

'=============================================================================
'  TICKER / SEGMENT VOLATILITY (nav code VT) - sheet "Tickers Volatility"
'-----------------------------------------------------------------------------
'  v3 (2026-10-10): rebuilt after the terminal-style "return history" screen.
'  One subject at a time, either
'     TICKER <GO>   one stock (any Yahoo symbol, bare TW codes try .TW/.TWO)
'     SEGMENT <GO>  a moomoo US plate (IndustryMap) or a TW group (tblGroups),
'                   built into a chain-linked composite index
'  PERIOD <GO> = how many bars the two left-hand charts show,
'  FREQ <GO>   = D / W / M bars.  Everything on the right is computed on the
'  whole sample (daily 520 bars, weekly 260, monthly 60) except the SD, which
'  is the SD of the PERIOD window (same as the screen this was modelled on).
'
'  Segment composite: US plates are market-cap weighted (IndustryMap column
'  5), TW groups have no cap column and are equal weighted.  At most
'  MAXMEM (60) of the largest members are fetched - every member is one
'  Yahoo request.
'
'  Page layout (page coordinates; modNav adds the bar afterwards):
'    A1 title.  Inputs: B2 ticker, E2 segment, B3 period, E3 freq.
'    rows 5-6 name / last / session / beta / segment
'    row 8 RETURN HISTORY summary, row 10 distribution title
'    charts from row 11 in A:J (built in VtFinish, AFTER NavAdd, because
'    series references do not follow a row/column insert)
'    right-hand tables from L8 on, constituents table from row 52
'    chart data lives from column DC (BH) - far away, black on black
'=============================================================================

Private Const VT_SHEET      As String = "Tickers Volatility"
Private Const IND_MAP       As String = "IndustryMap"
Private Const DC            As Long = 60          ' first chart-data column
Private Const MAXN          As Long = 1300        ' daily bars kept (~5y)
Private Const MAXMEM        As Long = 60
Private Const NBIN          As Long = 41

Private Const VT_TXT        As Long = 14540253    ' RGB(221,221,221)
Private Const VT_DIM        As Long = 9211020     ' RGB(140,140,140)
Private Const VT_UP         As Long = 3947980     ' RGB(204,60,60)  - TW: up is red
Private Const VT_DOWN       As Long = 7126016     ' RGB(0,190,108)  - TW: down is green
Private Const VT_OR         As Long = 1330687     ' RGB(255,150,20)
Private Const VT_BAND       As Long = 921870      ' RGB(14,14,14)

' ---- loaded subject (cached between PERIOD / FREQ changes) ----
Private gKey As String
Private gReady As Boolean
Private gMode As String                 ' T / S
Private sD() As Long, sP() As Double, sN As Long          ' subject, daily
Private bD() As Long, bP() As Double, bN As Long          ' benchmark, daily
Private gPeerOK As Boolean
Private gPeerDict As Object                                ' date -> peer return
Private gName As String, gSegName As String, gPeerLbl As String, gBenchLbl As String
Private gMktTW As Boolean
Private gSegInfo As String
Private mTk() As String, mW() As Double, mLast() As Double, mSd() As Double, mCnt As Long

' ---- geometry handed from the draw routine to VtFinish ----
Private gW As Long
Private gHistTitle As String
Private gDailyTitle As String
Private gHasMem As Boolean

'=============================================================================
Sub UpdateVolatilityAnalysis_Pro(Optional ByVal which As String = "")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(VT_SHEET)
    gReady = False

    Dim tk As String, sg As String, per As Long, fq As String, v As Variant
    tk = UCase$(Trim$(CStr(ws.Range("B2").Value)))
    sg = Trim$(CStr(ws.Range("E2").Value))
    v = ws.Range("B3").Value
    If IsNumeric(v) And Not IsEmpty(v) Then per = CLng(v)
    If per < 5 Then per = 180
    If per > 1200 Then per = 1200
    fq = UCase$(Left$(Trim$(CStr(ws.Range("E3").Value)), 1))
    If fq <> "W" And fq <> "M" Then fq = "D"

    Dim mode As String
    If which = "S" And sg <> "" Then
        mode = "S": tk = ""
    ElseIf which = "T" And tk <> "" Then
        mode = "T": sg = ""
    ElseIf tk <> "" Then
        mode = "T"
    ElseIf sg <> "" Then
        mode = "S"
    Else
        Call NavNotify("VT: type a ticker (B) or a segment (E)", True)
        Exit Sub
    End If

    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fin

    Dim segName As String
    If mode = "S" Then
        segName = ResolveSeg(sg)
        If segName = "" Then
            Application.EnableEvents = prevEv: Application.ScreenUpdating = True
            Exit Sub
        End If
    End If

    Dim key As String
    If mode = "T" Then key = "T|" & tk Else key = "S|" & segName
    If key <> gKey Or sN < 30 Then
        gKey = ""
        If mode = "T" Then
            If Not LoadTicker(tk) Then GoTo Fin
        Else
            If Not LoadSegment(segName) Then GoTo Fin
        End If
        gKey = key
    End If
    gMode = mode

    Call DrawPage(ws, mode, tk, segName, per, fq)

Fin:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    If Err.Number <> 0 Then
        Call NavNotify("VT error " & Err.Number & ": " & Err.Description, True)
    ElseIf gReady Then
        Call NavNotify("VT done - " & IIf(mode = "T", tk, segName) & "  " & per & " " & fq)
    End If
End Sub

'=============================================================================
'  Data loading
'=============================================================================
Private Function IsTwSym(ByVal t As String) As Boolean
    t = UCase$(Trim$(t))
    If Right$(t, 3) = ".TW" Or Right$(t, 4) = ".TWO" Then IsTwSym = True: Exit Function
    If InStr(t, ".") = 0 And Len(t) >= 4 And IsNumeric(Left$(t, 4)) Then IsTwSym = True
End Function

Private Function YSym(ByVal s As String) As String
    s = UCase$(Trim$(s))
    If Right$(s, 3) = ".TW" Or Right$(s, 4) = ".TWO" Or Right$(s, 3) = ".SS" Or Right$(s, 3) = ".SZ" Then
        YSym = s
    Else
        YSym = Replace(s, ".", "-")
    End If
End Function

Private Function VtFetch(ByVal sym As String, ByRef d() As Long, ByRef p() As Double) As Long
    Dim dd() As Date, pp() As Double, s As String, i As Long, n As Long, st As Long
    s = sym
    If Not GetHistoricalData(s, dd, pp) Then Exit Function
    n = UBound(pp) + 1
    If n < 5 Then Exit Function
    If n > MAXN Then st = n - MAXN
    ReDim d(0 To n - st - 1): ReDim p(0 To n - st - 1)
    For i = st To n - 1
        d(i - st) = Int(CDbl(dd(i)))
        p(i - st) = pp(i)
    Next i
    VtFetch = n - st
End Function

Private Function LoadBench(ByVal isTW As Boolean) As Boolean
    Dim sym As String
    If isTW Then sym = "%5ETWII": gBenchLbl = "^TWII" Else sym = "SPY": gBenchLbl = "SPY"
    bN = VtFetch(sym, bD, bP)
    LoadBench = (bN > 30)
End Function

Private Function LoadTicker(ByVal tk As String) As Boolean
    mCnt = 0: gHasMem = False: gPeerOK = False: gSegName = "-": gPeerLbl = "": gSegInfo = ""
    Application.StatusBar = "VT: fetching " & tk & " ..."
    sN = VtFetch(tk, sD, sP)
    If sN < 30 Then
        sN = 0
        Call NavNotify("VT: could not download enough price history for " & tk, True)
        Exit Function
    End If
    gMktTW = IsTwSym(tk)
    If Not LoadBench(gMktTW) Then bN = 0
    On Error Resume Next
    Dim t2 As String: t2 = tk
    gName = GetCompanyName(t2)
    On Error GoTo 0
    If gName = "" Then gName = tk

    ' segment + peers
    Dim seg As String, syms() As String, caps() As Double, m As Long, total As Long
    seg = SegOfTicker(tk, gMktTW)
    If seg <> "" Then
        gSegName = seg
        m = SegMembers(seg, gMktTW, syms, caps, total)
        If m > 0 Then Call BuildPeers(tk, syms, caps, m)
    End If
    LoadTicker = True
End Function

' top 3 other members -> equal-weight composite, kept as date -> return
Private Sub BuildPeers(ByVal tk As String, ByRef syms() As String, ByRef caps() As Double, ByVal m As Long)
    Dim ps(0 To 2) As String, pw(0 To 2) As Double, np As Long, i As Long
    Dim me1 As String: me1 = NormTk(tk)
    For i = 0 To m - 1
        If NormTk(syms(i)) <> me1 Then
            ps(np) = syms(i): pw(np) = 1: np = np + 1
            If np = 3 Then Exit For
        End If
    Next i
    If np = 0 Or bN < 30 Then Exit Sub
    Dim od() As Long, op() As Double, on1 As Long
    Dim lbl As String, k As Long
    Dim tks() As String, ws2() As Double, ml() As Double, ms() As Double, mc As Long
    ReDim tks(0 To np - 1): ReDim ws2(0 To np - 1)
    For k = 0 To np - 1: tks(k) = ps(k): ws2(k) = pw(k): Next k
    on1 = Composite(tks, ws2, np, od, op, ml, ms, mc, "peers")
    If on1 < 30 Then Exit Sub
    Dim dd() As Long, rr() As Double, dm As Long
    dm = DailyRets(od, op, on1, dd, rr)
    Set gPeerDict = CreateObject("Scripting.Dictionary")
    For k = 0 To dm - 1: gPeerDict(dd(k)) = rr(k): Next k
    For k = 0 To np - 1
        lbl = lbl & IIf(lbl = "", "", ", ") & NormTk(ps(k))
    Next k
    gPeerLbl = lbl
    gPeerOK = True
End Sub

Private Function LoadSegment(ByVal segName As String) As Boolean
    mCnt = 0: gHasMem = False: gPeerOK = False: gPeerLbl = ""
    gName = segName: gSegName = segName
    Dim mkt As String: mkt = SegMarket(segName)
    gMktTW = (mkt = "TW")
    If Not LoadBench(gMktTW) Then
        Call NavNotify("VT: could not load the benchmark index", True)
        Exit Function
    End If
    Dim syms() As String, caps() As Double, m As Long, total As Long
    m = SegMembers(segName, gMktTW, syms, caps, total)
    If m = 0 Then
        Call NavNotify("VT: no constituents found for " & segName, True)
        Exit Function
    End If
    Dim od() As Long, op() As Double, on1 As Long
    Dim ml() As Double, ms() As Double, mc As Long
    on1 = Composite(syms, caps, m, od, op, ml, ms, mc, segName)
    If on1 < 30 Then
        Call NavNotify("VT: not enough constituent price history for " & segName, True)
        Exit Function
    End If
    sN = on1
    ReDim sD(0 To sN - 1): ReDim sP(0 To sN - 1)
    Dim i As Long
    For i = 0 To sN - 1: sD(i) = od(i): sP(i) = op(i): Next i
    mCnt = 0
    ReDim mTk(0 To m): ReDim mW(0 To m): ReDim mLast(0 To m): ReDim mSd(0 To m)
    Dim wsum As Double
    For i = 0 To m - 1
        If ms(i) > 0 Or ml(i) <> 0 Then
            mTk(mCnt) = NormTk(syms(i)): mW(mCnt) = caps(i): mLast(mCnt) = ml(i): mSd(mCnt) = ms(i)
            wsum = wsum + caps(i): mCnt = mCnt + 1
        End If
    Next i
    If wsum > 0 Then
        For i = 0 To mCnt - 1: mW(i) = mW(i) / wsum: Next i
    End If
    gHasMem = (mCnt > 0)
    gSegInfo = mc & " of " & total & IIf(total > m, " (top " & m & " by cap)", "") & IIf(gMktTW, " . equal weight", " . cap weight")
    LoadSegment = True
End Function

Private Function NormTk(ByVal s As String) As String
    s = UCase$(Trim$(s))
    If Right$(s, 4) = ".TWO" Then s = Left$(s, Len(s) - 4)
    If Right$(s, 3) = ".TW" Then s = Left$(s, Len(s) - 3)
    NormTk = s
End Function

' Chain-linked composite on the benchmark's day grid. Weights w(); returns the
' number of points; out arrays od/op (index from 100); per-member last return
' and daily SD in ml()/ms() (aligned with syms order; 0 when no data).
Private Function Composite(ByRef syms() As String, ByRef w() As Double, ByVal m As Long, _
                           ByRef od() As Long, ByRef op() As Double, _
                           ByRef ml() As Double, ByRef ms() As Double, ByRef nOk As Long, _
                           ByVal lbl As String) As Long
    Dim nb As Long: nb = bN
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim t As Long, j As Long, k As Long
    For t = 0 To nb - 1: idx(bD(t)) = t: Next t
    Dim sumW() As Double, sumR() As Double
    ReDim sumW(0 To nb - 1): ReDim sumR(0 To nb - 1)
    ReDim ml(0 To m): ReDim ms(0 To m)
    Dim d() As Long, p() As Double, n As Long, r As Double
    Dim q() As Double, qn As Long, mu As Double, s2 As Double, t0 As Double
    t0 = Timer
    nOk = 0
    For j = 0 To m - 1
        If w(j) > 0 Then
            Application.StatusBar = "VT: " & lbl & "  member " & (j + 1) & "/" & m & " " & syms(j) & _
                                    "  (" & Format(Timer - t0, "0") & "s)"
            DoEvents
            n = VtFetch(YSym(syms(j)), d, p)
            If n >= 5 Then
                nOk = nOk + 1
                ReDim q(0 To n - 2): qn = n - 1: mu = 0
                For k = 1 To n - 1
                    If p(k - 1) > 0 Then r = p(k) / p(k - 1) - 1 Else r = 0
                    q(k - 1) = r: mu = mu + r
                    If idx.Exists(d(k)) Then
                        t = idx(d(k))
                        sumW(t) = sumW(t) + w(j)
                        sumR(t) = sumR(t) + w(j) * r
                    End If
                Next k
                mu = mu / qn: s2 = 0
                For k = 0 To qn - 1: s2 = s2 + (q(k) - mu) ^ 2: Next k
                If qn > 1 Then ms(j) = Sqr(s2 / (qn - 1))
                ml(j) = q(qn - 1)
            End If
        End If
    Next j
    ReDim od(0 To nb - 1): ReDim op(0 To nb - 1)
    Dim ix As Double: ix = 100
    Dim started As Boolean, np As Long
    For t = 0 To nb - 1
        If Not started Then
            If sumW(t) > 0 Then
                started = True
                od(np) = bD(t): op(np) = ix: np = np + 1
            End If
        Else
            If sumW(t) > 0 Then ix = ix * (1 + sumR(t) / sumW(t))
            od(np) = bD(t): op(np) = ix: np = np + 1
        End If
    Next t
    Composite = np
End Function

'=============================================================================
'  Segment lists (IndustryMap = moomoo US plates, tblGroups = TW groups)
'=============================================================================
Private Function MapData() As Variant
    Dim wm As Worksheet
    On Error Resume Next
    Set wm = ThisWorkbook.Sheets(IND_MAP)
    On Error GoTo 0
    If wm Is Nothing Then Exit Function
    Dim lastR As Long: lastR = wm.Cells(wm.Rows.count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function
    MapData = wm.Range(wm.Cells(2, 1), wm.Cells(lastR, 6)).Value
End Function

Private Function SegKey(ByVal s As String) As String
    s = Replace(Replace(s, ChrW(&H3000), ""), " ", "")
    SegKey = LCase$(s)
End Function

' all segment names: TW groups first, then US plates
Private Function SegNames() As Variant
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare
    Dim a As Variant, i As Long
    On Error Resume Next
    a = Sanner.GetGroupNames("TW")
    On Error GoTo 0
    If Not IsEmpty(a) Then
        For i = LBound(a) To UBound(a)
            If Not dic.Exists(CStr(a(i))) Then dic.Add CStr(a(i)), "TW"
        Next i
    End If
    Dim v As Variant: v = MapData()
    If Not IsEmpty(v) Then
        Dim r As Long, nm As String
        For r = 1 To UBound(v, 1)
            nm = Trim$(CStr(v(r, 2)))
            If nm <> "" Then
                If Not dic.Exists(nm) Then dic.Add nm, "US"
            End If
        Next r
    End If
    If dic.count > 0 Then SegNames = dic.keys
End Function

Private Function SegMarket(ByVal nm As String) As String
    Dim a As Variant, i As Long
    On Error Resume Next
    a = Sanner.GetGroupNames("TW")
    On Error GoTo 0
    If Not IsEmpty(a) Then
        For i = LBound(a) To UBound(a)
            If SegKey(CStr(a(i))) = SegKey(nm) Then SegMarket = "TW": Exit Function
        Next i
    End If
    SegMarket = "US"
End Function

' typed text -> exact segment name; "" (with a status-bar hint) if ambiguous
Private Function ResolveSeg(ByVal txt As String) As String
    Dim a As Variant: a = SegNames()
    If IsEmpty(a) Then
        Call NavNotify("VT: no segment lists loaded (run IMAP / fill tblGroups)", True)
        Exit Function
    End If
    Dim k As String: k = SegKey(txt)
    Dim i As Long, hit As String, nHit As Long, cand As String
    For i = LBound(a) To UBound(a)
        If SegKey(CStr(a(i))) = k Then ResolveSeg = CStr(a(i)): Exit Function
    Next i
    For i = LBound(a) To UBound(a)
        If InStr(SegKey(CStr(a(i))), k) > 0 Then
            nHit = nHit + 1
            If nHit = 1 Then hit = CStr(a(i))
            If nHit <= 5 Then cand = cand & IIf(cand = "", "", " | ") & CStr(a(i))
        End If
    Next i
    If nHit = 1 Then
        ResolveSeg = hit
    ElseIf nHit > 1 Then
        Call NavNotify("VT: '" & txt & "' matches " & nHit & ": " & cand, True)
    Else
        Call NavNotify("VT: no segment named '" & txt & "'", True)
    End If
End Function

Private Function SegOfTicker(ByVal tk As String, ByVal isTW As Boolean) As String
    Dim nt As String: nt = NormTk(tk)
    Dim i As Long, j As Long
    If isTW Then
        Dim a As Variant, mem As Variant
        On Error Resume Next
        a = Sanner.GetGroupNames("TW")
        On Error GoTo 0
        If IsEmpty(a) Then Exit Function
        For i = LBound(a) To UBound(a)
            mem = Empty
            On Error Resume Next
            mem = Sanner.GetSectorTickers("TW", CStr(a(i)))
            On Error GoTo 0
            If Not IsEmpty(mem) Then
                For j = LBound(mem) To UBound(mem)
                    If NormTk(CStr(mem(j))) = nt Then SegOfTicker = CStr(a(i)): Exit Function
                Next j
            End If
        Next i
    Else
        Dim v As Variant: v = MapData()
        If IsEmpty(v) Then Exit Function
        Dim r As Long, s As String
        For r = 1 To UBound(v, 1)
            s = UCase$(Trim$(CStr(v(r, 3))))
            If s = nt Or Replace(s, ".", "-") = nt Then SegOfTicker = Trim$(CStr(v(r, 2))): Exit Function
        Next r
    End If
End Function

' members sorted (US by cap desc), capped to MAXMEM; total = uncapped count
Private Function SegMembers(ByVal nm As String, ByVal isTW As Boolean, ByRef syms() As String, _
                            ByRef caps() As Double, ByRef total As Long) As Long
    Dim n As Long, i As Long, j As Long
    If isTW Then
        Dim mem As Variant
        On Error Resume Next
        mem = Sanner.GetSectorTickers("TW", nm)
        On Error GoTo 0
        If IsEmpty(mem) Then Exit Function
        total = UBound(mem) - LBound(mem) + 1
        n = total: If n > MAXMEM Then n = MAXMEM
        ReDim syms(0 To n - 1): ReDim caps(0 To n - 1)
        For i = 0 To n - 1
            syms(i) = CStr(mem(LBound(mem) + i)): caps(i) = 1
        Next i
        SegMembers = n
    Else
        Dim v As Variant: v = MapData()
        If IsEmpty(v) Then Exit Function
        Dim r As Long, cnt As Long
        Dim ts() As String, tc() As Double
        ReDim ts(0 To UBound(v, 1)): ReDim tc(0 To UBound(v, 1))
        For r = 1 To UBound(v, 1)
            If SegKey(CStr(v(r, 2))) = SegKey(nm) Then
                If Trim$(CStr(v(r, 3))) <> "" Then
                    ts(cnt) = Trim$(CStr(v(r, 3)))
                    If IsNumeric(v(r, 5)) And Not IsEmpty(v(r, 5)) Then tc(cnt) = CDbl(v(r, 5)) Else tc(cnt) = 0
                    cnt = cnt + 1
                End If
            End If
        Next r
        If cnt = 0 Then Exit Function
        total = cnt
        Dim tmpS As String, tmpC As Double
        For i = 1 To cnt - 1
            tmpS = ts(i): tmpC = tc(i): j = i - 1
            Do While j >= 0
                If tc(j) >= tmpC Then Exit Do
                ts(j + 1) = ts(j): tc(j + 1) = tc(j): j = j - 1
            Loop
            ts(j + 1) = tmpS: tc(j + 1) = tmpC
        Next i
        n = cnt: If n > MAXMEM Then n = MAXMEM
        ReDim syms(0 To n - 1): ReDim caps(0 To n - 1)
        For i = 0 To n - 1
            syms(i) = ts(i)
            caps(i) = tc(i): If caps(i) <= 0 Then caps(i) = 1
        Next i
        SegMembers = n
    End If
End Function

'=============================================================================
'  Series helpers
'=============================================================================
Private Function DailyRets(ByRef d() As Long, ByRef p() As Double, ByVal n As Long, _
                           ByRef rd() As Long, ByRef rr() As Double) As Long
    Dim i As Long, m As Long
    If n < 2 Then Exit Function
    ReDim rd(0 To n - 2): ReDim rr(0 To n - 2)
    For i = 1 To n - 1
        If p(i - 1) > 0 Then
            rd(m) = d(i): rr(m) = p(i) / p(i - 1) - 1: m = m + 1
        End If
    Next i
    DailyRets = m
End Function

Private Sub Resample(ByRef d() As Long, ByRef p() As Double, ByVal n As Long, ByVal f As String, _
                     ByRef od() As Long, ByRef op() As Double, ByRef m As Long)
    Dim i As Long, k As Long, pk As Long
    ReDim od(0 To n - 1): ReDim op(0 To n - 1)
    m = 0
    If f = "D" Then
        For i = 0 To n - 1: od(i) = d(i): op(i) = p(i): Next i
        m = n: Exit Sub
    End If
    pk = -1
    For i = 0 To n - 1
        If f = "W" Then
            k = Int((d(i) - 2) / 7)
        Else
            k = Year(CDate(d(i))) * 12 + Month(CDate(d(i)))
        End If
        If i > 0 And k <> pk Then
            od(m) = d(i - 1): op(m) = p(i - 1): m = m + 1
        End If
        pk = k
    Next i
    od(m) = d(n - 1): op(m) = p(n - 1): m = m + 1
End Sub

Private Function MeanS(ByRef a() As Double, ByVal i0 As Long, ByVal i1 As Long) As Double
    Dim i As Long, s As Double
    If i1 < i0 Then Exit Function
    For i = i0 To i1: s = s + a(i): Next i
    MeanS = s / (i1 - i0 + 1)
End Function

Private Function SdS(ByRef a() As Double, ByVal i0 As Long, ByVal i1 As Long) As Double
    Dim n As Long, i As Long, m As Double, s As Double
    n = i1 - i0 + 1
    If n < 2 Then Exit Function
    m = MeanS(a, i0, i1)
    For i = i0 To i1: s = s + (a(i) - m) ^ 2: Next i
    SdS = Sqr(s / (n - 1))
End Function

Private Function NormCdf(ByVal z As Double) As Double
    NormCdf = Application.WorksheetFunction.Norm_S_Dist(z, True)
End Function

Private Function NiceBin(ByVal target As Double) As Double
    Dim c As Variant, i As Long, best As Double, bd As Double, dd As Double
    c = Array(0.0005, 0.001, 0.002, 0.0025, 0.005, 0.01, 0.02, 0.025, 0.05, 0.1)
    bd = 1E+99
    If target <= 0 Then NiceBin = 0.005: Exit Function
    For i = 0 To UBound(c)
        dd = Abs(Log(c(i) / target))
        If dd < bd Then bd = dd: best = c(i)
    Next i
    NiceBin = best
End Function

Private Function PctTxt(ByVal x As Double, Optional ByVal dec As Long = 2, Optional ByVal signed As Boolean = True) As String
    Dim f As String: f = "0." & String(dec, "0")
    If signed Then
        PctTxt = IIf(x >= 0, "+", "-") & Format(Abs(x) * 100, f) & "%"
    Else
        PctTxt = Format(x * 100, f) & "%"
    End If
End Function

Private Function ColOf(ByVal x As Double) As Long
    If x > 0 Then
        ColOf = VT_UP
    ElseIf x < 0 Then
        ColOf = VT_DOWN
    Else
        ColOf = VT_TXT
    End If
End Function

'=============================================================================
'  Page drawing
'=============================================================================
Private Sub PutC(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal v As Variant, _
                Optional ByVal fmt As String = "", Optional ByVal col As Long = -1, _
                Optional ByVal bold As Boolean = False, Optional ByVal al As Long = xlRight)
    With ws.Cells(r, c)
        If fmt = "@" Then .NumberFormat = "@"
        .Value = v
        If fmt <> "" And fmt <> "@" Then .NumberFormat = fmt
        If col <> -1 Then .Font.Color = col
        If bold Then .Font.Bold = True
        .HorizontalAlignment = al
    End With
End Sub

Private Sub Hdr(ws As Worksheet, ByVal r As Long, ByVal c As Long, ParamArray labs() As Variant)
    Dim i As Long
    For i = 0 To UBound(labs)
        With ws.Cells(r, c + i)
            .Value = labs(i)
            .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 8
            .HorizontalAlignment = IIf(i = 0, xlLeft, xlRight)
        End With
    Next i
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + UBound(labs))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RR4_LINE
    End With
End Sub

Private Sub Ttl(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal t As String, Optional ByVal sub1 As String = "", Optional ByVal subCol As Long = 0)
    With ws.Cells(r, c)
        .Value = t: .Font.Bold = True: .Font.Color = RR4_ACCENT: .HorizontalAlignment = xlLeft
    End With
    If sub1 <> "" Then
        With ws.Cells(r, IIf(subCol > 0, subCol, c + 2))
            .Value = sub1: .Font.Color = VT_DIM: .Font.Size = 8: .HorizontalAlignment = xlLeft
        End With
    End If
End Sub

Private Sub InputCell(ws As Worksheet, ByVal addr As String, ByVal v As Variant, Optional ByVal span As Long = 1)
    Dim rg As Range: Set rg = ws.Range(addr).Resize(1, span)
    rg.Interior.Color = RR4_INPUT_BG
    rg.Font.Color = RR4_INPUT_FG: rg.Font.Bold = True
    With ws.Range(addr)
        .NumberFormat = "@"
        .Value = v
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub ResetPage(ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.ChartObjects().count To 1 Step -1: ws.ChartObjects(i).Delete: Next i
    On Error GoTo 0
    ws.Cells.Clear
    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = VT_TXT
        .Font.Name = "Consolas": .Font.Size = 9
        .VerticalAlignment = xlCenter: .HorizontalAlignment = xlRight
        .RowHeight = 17
    End With
    ws.Rows(1).RowHeight = 24
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B:J").ColumnWidth = 11
    ws.Columns("K").ColumnWidth = 3
    ws.Columns("L").ColumnWidth = 16
    ws.Columns("M:S").ColumnWidth = 11
    With ws.Cells(1, 1)
        .Value = "TICKER / SEGMENT VOLATILITY"
        .HorizontalAlignment = xlLeft: .Font.Size = 14: .Font.Bold = True: .Font.Color = RR4_ACCENT
    End With
End Sub

Private Sub DrawPage(ws As Worksheet, ByVal mode As String, ByVal tk As String, ByVal segName As String, _
                     ByVal per As Long, ByVal fq As String)
    Dim i As Long, j As Long
    Dim fqName As String, unit As String
    Select Case fq
        Case "W": fqName = "weekly": unit = "weeks"
        Case "M": fqName = "monthly": unit = "months"
        Case Else: fqName = "daily": unit = "sessions"
    End Select

    ' ---- sample at the chosen frequency ----
    Dim rd() As Long, rp() As Double, rn As Long
    Call Resample(sD, sP, sN, fq, rd, rp, rn)
    Dim smp As Long
    Select Case fq
        Case "W": smp = 260
        Case "M": smp = 60
        Case Else: smp = 520
    End Select
    Dim st As Long: st = 0
    If rn > smp + 1 Then st = rn - smp - 1
    Dim m As Long: m = rn - 1 - st
    If m < 12 Then
        Call NavNotify("VT: not enough history at this frequency (" & m & " returns)", True)
        Exit Sub
    End If
    Dim rr() As Double, dt() As Long, px() As Double
    ReDim rr(0 To m - 1): ReDim dt(0 To m - 1): ReDim px(0 To m)
    For i = 0 To m
        px(i) = rp(st + i)
    Next i
    For i = 0 To m - 1
        If px(i) > 0 Then rr(i) = px(i + 1) / px(i) - 1
        dt(i) = rd(st + i + 1)
    Next i

    Dim w As Long: w = per: If w > m Then w = m
    Dim w0 As Long: w0 = m - w
    Dim sdW As Double, muW As Double, muF As Double
    sdW = SdS(rr, w0, m - 1): muW = MeanS(rr, w0, m - 1): muF = MeanS(rr, 0, m - 1)
    If sdW <= 0 Then Call NavNotify("VT: zero variance in the window", True): Exit Sub
    Dim lastR As Double: lastR = rr(m - 1)
    Dim prevPx As Double: prevPx = px(m - 1)
    Dim lastPx As Double: lastPx = px(m)

    Call ResetPage(ws)
    ws.Columns("A").ColumnWidth = 14

    ' ---- inputs ----
    PutC ws, 2, 1, "TICKER <GO>", , VT_DIM, , xlLeft
    Call InputCell(ws, "B2", IIf(mode = "T", tk, ""), 1)
    PutC ws, 2, 4, "SEGMENT <GO>", , VT_DIM, , xlLeft
    Call InputCell(ws, "E2", IIf(mode = "S", segName, ""), 4)
    PutC ws, 3, 1, "PERIOD <GO>", , VT_DIM, , xlLeft
    Call InputCell(ws, "B3", per, 1)
    PutC ws, 3, 4, "FREQ <GO>", , VT_DIM, , xlLeft
    Call InputCell(ws, "E3", fq, 1)
    Call BuildSegList(ws)
    On Error Resume Next
    With ws.Range("E3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="D,W,M"
        .ShowError = False
    End With
    With ws.Range("E2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="=VtSegList"
        .ShowError = False
    End With
    On Error GoTo 0

    ' ---- identity block ----
    Dim lastD As Double: lastD = sP(sN - 1)
    Dim dayR As Double
    If sN >= 2 And sP(sN - 2) > 0 Then dayR = sP(sN - 1) / sP(sN - 2) - 1
    Dim dRd() As Long, dRr() As Double, dm As Long
    dm = DailyRets(sD, sP, sN, dRd, dRr)

    PutC ws, 5, 1, "NAME", , VT_DIM, , xlLeft
    PutC ws, 5, 2, gName, , RR4_INPUT_FG, True, xlLeft
    PutC ws, 5, 8, "LAST PRICE", , VT_DIM, , xlRight
    PutC ws, 5, 9, lastD, "#,##0.00", VT_TXT, True
    PutC ws, 5, 10, dayR, "+0.00%;-0.00%", ColOf(dayR), True
    PutC ws, 6, 1, "SESSION", , VT_DIM, , xlLeft
    PutC ws, 6, 2, Format(CDate(sD(sN - 1)), "ddd d mmm"), "@", RR4_INPUT_FG, False, xlLeft
    PutC ws, 6, 4, "BETA", , VT_DIM, , xlLeft
    PutC ws, 6, 8, "SEGMENT", , VT_DIM, , xlRight

    ' ---- proxy numbers (daily, full sample) ----
    Dim bDict As Object: Set bDict = CreateObject("Scripting.Dictionary")
    Dim bRd() As Long, bRr() As Double, bm As Long
    bm = DailyRets(bD, bP, bN, bRd, bRr)
    For i = 0 To bm - 1: bDict(bRd(i)) = bRr(i): Next i
    Dim c1 As Double, c60 As Double, beta1 As Double, sameD As Double, nA As Long
    Dim haveB As Boolean
    haveB = ProxyStats(dRd, dRr, dm, bDict, c1, c60, beta1, sameD, nA)
    Dim c1p As Double, c60p As Double, beta1p As Double, sameDp As Double, nAp As Long, haveP As Boolean
    If gPeerOK And mode = "T" Then haveP = ProxyStats(dRd, dRr, dm, gPeerDict, c1p, c60p, beta1p, sameDp, nAp)
    If haveB Then
        PutC ws, 6, 5, beta1, "0.00", VT_TXT, True
    Else
        PutC ws, 6, 5, "-", , VT_DIM
    End If
    PutC ws, 6, 9, IIf(mode = "S", gSegInfo, gSegName), , VT_TXT, False, xlLeft
    ws.Cells(6, 9).HorizontalAlignment = xlLeft
    ws.Range("A7:S7").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("A7:S7").Borders(xlEdgeBottom).Color = RR4_LINE

    ' ---- summary line ----
    Dim nBeyond As Long
    For i = w0 To m - 1
        If Abs(rr(i) - muW) > 2 * sdW Then nBeyond = nBeyond + 1
    Next i
    Ttl ws, 8, 1, "RETURN HISTORY"
    With ws.Cells(8, 3)
        .Value = fqName & "  .  last " & w & " " & unit & "  .  " & Format(CDate(dt(w0)), "yyyy-mm-dd") & " to " & Format(CDate(dt(m - 1)), "yyyy-mm-dd") & _
                 "  .  SD " & PctTxt(sdW, 2, False) & "  .  " & nBeyond & " of " & w & " past 2 SD  .  last " & PctTxt(lastR)
        .Font.Color = VT_DIM: .Font.Size = 8: .HorizontalAlignment = xlLeft
    End With
    If w < per Then
        PutC ws, 9, 1, "only " & w & " " & unit & " available - PERIOD capped", , VT_OR, False, xlLeft
    End If

    ' ---- histogram data (page cols DC..) ----
    Dim bw As Double: bw = NiceBin(sdW / 4)
    Dim cnt() As Double, expd() As Double, j2 As Long
    ReDim cnt(0 To NBIN - 1): ReDim expd(0 To NBIN - 1)
    Dim lastBin As Long
    For i = w0 To m - 1
        j2 = BinOf(rr(i), bw)
        cnt(j2) = cnt(j2) + 1
    Next i
    lastBin = BinOf(lastR, bw)
    Dim lo As Double, hi As Double
    For j = 0 To NBIN - 1
        If j = 0 Then
            lo = -1E+99
        Else
            lo = (j - 20.5) * bw
        End If
        If j = NBIN - 1 Then
            hi = 1E+99
        Else
            hi = (j - 19.5) * bw
        End If
        Dim pl As Double, ph As Double
        If lo < -1E+98 Then pl = 0 Else pl = NormCdf((lo - muW) / sdW)
        If hi > 1E+98 Then ph = 1 Else ph = NormCdf((hi - muW) / sdW)
        expd(j) = w * (ph - pl)
    Next j
    PutC ws, 1, DC, "BIN", , VT_DIM, , xlLeft
    PutC ws, 1, DC + 1, "COUNT", , VT_DIM
    PutC ws, 1, DC + 2, "NORMAL", , VT_DIM
    PutC ws, 1, DC + 3, "LAST", , VT_DIM
    PutC ws, 1, DC + 4, "SIDE", , VT_DIM
    For j = 0 To NBIN - 1
        PutC ws, 2 + j, DC, Format((j - 20) * bw * 100, "0.0#") & "%", "@", VT_DIM, , xlLeft
        PutC ws, 2 + j, DC + 1, cnt(j), "0"
        PutC ws, 2 + j, DC + 2, expd(j), "0.00"
        If j = lastBin Then PutC ws, 2 + j, DC + 3, cnt(j), "0"
        PutC ws, 2 + j, DC + 4, IIf(j > 20, 1, IIf(j < 20, -1, 0)), "0"
    Next j

    ' ---- daily window data ----
    PutC ws, 1, DC + 6, "DATE", , VT_DIM, , xlLeft
    PutC ws, 1, DC + 7, "UP", , VT_DIM
    PutC ws, 1, DC + 8, "DOWN", , VT_DIM
    PutC ws, 1, DC + 9, "BEYOND", , VT_DIM
    PutC ws, 1, DC + 10, "+1SD", , VT_DIM
    PutC ws, 1, DC + 11, "+2SD", , VT_DIM
    PutC ws, 1, DC + 12, "-1SD", , VT_DIM
    PutC ws, 1, DC + 13, "-2SD", , VT_DIM
    Dim fmtD As String: fmtD = IIf(fq = "M", "mmm yy", "dd mmm yy")
    For i = 0 To w - 1
        PutC ws, 2 + i, DC + 6, Format(CDate(dt(w0 + i)), IIf(fq = "M", "mmm yy", "dd mmm")), "@", VT_DIM, , xlLeft
        If Abs(rr(w0 + i) - muW) > 2 * sdW Then
            PutC ws, 2 + i, DC + 9, rr(w0 + i), "0.00%"
        ElseIf rr(w0 + i) >= 0 Then
            PutC ws, 2 + i, DC + 7, rr(w0 + i), "0.00%"
        Else
            PutC ws, 2 + i, DC + 8, rr(w0 + i), "0.00%"
        End If
        PutC ws, 2 + i, DC + 10, muW + sdW, "0.00%"
        PutC ws, 2 + i, DC + 11, muW + 2 * sdW, "0.00%"
        PutC ws, 2 + i, DC + 12, muW - sdW, "0.00%"
        PutC ws, 2 + i, DC + 13, muW - 2 * sdW, "0.00%"
    Next i
    gW = w
    gHistTitle = "WHERE THE LAST SESSION SITS  .  bin " & Format(bw * 100, "0.0#") & "%  .  -2SD " & PctTxt(muW - 2 * sdW) & _
                 "  -1SD " & PctTxt(muW - sdW) & "  +1SD " & PctTxt(muW + sdW) & "  +2SD " & PctTxt(muW + 2 * sdW)
    gDailyTitle = UCase$(fqName) & " RETURNS  .  dashed = +/-1 and +/-2 SD  .  orange = beyond 2 SD"

    ' =========================== right-hand tables ===========================
    Dim r As Long, k As Long
    ' --- standard deviation bands ---
    Ttl ws, 8, 12, "STANDARD DEVIATION / SD", "from the " & Format(prevPx, "#,##0.00") & " close", 14
    Hdr ws, 9, 12, "BAND", "MOVE", "LOW", "HIGH", "DAYS INSIDE"
    For k = 1 To 3
        r = 9 + k
        Dim ins As Long
        ins = 0
        For i = 0 To m - 1
            If Abs(rr(i)) <= k * sdW Then ins = ins + 1
        Next i
        PutC ws, r, 12, "+/-" & k & " SD", "@", VT_TXT, True, xlLeft
        PutC ws, r, 13, "+/-" & PctTxt(k * sdW, 2, False), "@", VT_TXT
        PutC ws, r, 14, prevPx * (1 - k * sdW), "#,##0.00", VT_DOWN
        PutC ws, r, 15, prevPx * (1 + k * sdW), "#,##0.00", VT_UP
        PutC ws, r, 16, ins / m, "0.0%", VT_TXT
    Next k

    ' --- proxy check ---
    Ttl ws, 14, 12, "PROXY CHECK", "daily moves vs market and segment  .  " & nA & " sessions", 14
    Hdr ws, 15, 12, "AGAINST", "CORRELATION", "60D CORR", "BETA", "SAME DIR"
    PutC ws, 16, 12, gBenchLbl, "@", VT_TXT, True, xlLeft
    If haveB Then
        PutC ws, 16, 13, c1, "0.00", VT_TXT
        PutC ws, 16, 14, c60, "0.00", VT_TXT
        PutC ws, 16, 15, beta1, "0.00", VT_TXT
        PutC ws, 16, 16, sameD, "0.0%", VT_TXT
    Else
        PutC ws, 16, 13, "-", , VT_DIM
    End If
    If haveP Then
        PutC ws, 17, 12, gPeerLbl, "@", VT_TXT, True, xlLeft
        PutC ws, 17, 13, c1p, "0.00", VT_TXT
        PutC ws, 17, 14, c60p, "0.00", VT_TXT
        PutC ws, 17, 15, beta1p, "0.00", VT_TXT
        PutC ws, 17, 16, sameDp, "0.0%", VT_TXT
    Else
        PutC ws, 17, 12, IIf(mode = "S", "segment vs market only", "no segment peers"), "@", VT_DIM, False, xlLeft
    End If

    ' --- up vs down ---
    Ttl ws, 19, 12, "UP DAYS VS DOWN DAYS", m & " " & unit, 14
    Hdr ws, 20, 12, "SIDE", "DAYS", "SHARE", "AVERAGE", "WEIGHTED"
    Dim nUp As Long, nDn As Long, sUp As Double, sDn As Double
    For i = 0 To m - 1
        If rr(i) > 0 Then
            nUp = nUp + 1: sUp = sUp + rr(i)
        ElseIf rr(i) < 0 Then
            nDn = nDn + 1: sDn = sDn + rr(i)
        End If
    Next i
    Dim aUp As Double, aDn As Double
    If nUp > 0 Then aUp = sUp / nUp
    If nDn > 0 Then aDn = sDn / nDn
    PutC ws, 21, 12, "Up", "@", VT_UP, True, xlLeft
    PutC ws, 21, 13, nUp, "0", VT_TXT
    PutC ws, 21, 14, nUp / m, "0.0%", VT_TXT
    PutC ws, 21, 15, aUp, "+0.00%;-0.00%", VT_UP
    PutC ws, 21, 16, aUp * nUp / m, "+0.00%;-0.00%", VT_UP
    PutC ws, 22, 12, "Down", "@", VT_DOWN, True, xlLeft
    PutC ws, 22, 13, nDn, "0", VT_TXT
    PutC ws, 22, 14, nDn / m, "0.0%", VT_TXT
    PutC ws, 22, 15, aDn, "+0.00%;-0.00%", VT_DOWN
    PutC ws, 22, 16, aDn * nDn / m, "+0.00%;-0.00%", VT_DOWN

    ' --- stretch ---
    Ttl ws, 24, 12, "STRETCH", "run so far against its own SD", 14
    Hdr ws, 25, 12, "WINDOW", "RUN", "IN SD"
    Dim lens As Variant: lens = Array(5, 20, 60)
    For k = 0 To 2
        Dim kk As Long: kk = lens(k)
        r = 26 + k
        PutC ws, r, 12, kk & " " & unit, "@", VT_TXT, True, xlLeft
        If m >= kk Then
            Dim cum As Double: cum = px(m) / px(m - kk) - 1
            PutC ws, r, 13, cum, "+0.00%;-0.00%", ColOf(cum)
            PutC ws, r, 14, cum / (sdW * Sqr(kk)), "+0.0"" SD"";-0.0"" SD""", ColOf(cum)
        Else
            PutC ws, r, 13, "-", , VT_DIM
        End If
    Next k
    Dim hiPx As Double, hiI As Long
    For i = 0 To m
        If px(i) >= hiPx Then hiPx = px(i): hiI = i
    Next i
    PutC ws, 29, 12, "FROM HIGH", "@", VT_TXT, True, xlLeft
    PutC ws, 29, 13, px(m) / hiPx - 1, "+0.00%;-0.00%", ColOf(px(m) / hiPx - 1)
    PutC ws, 29, 14, Format(CDate(rd(st + hiI)), "d mmm yy"), "@", VT_DIM

    ' --- return bands (0.75 SD steps around the sample mean) ---
    Ttl ws, 31, 12, "RETURN BANDS", fqName & "  .  " & m & " " & unit & "  .  0.75 SD steps", 14
    Hdr ws, 32, 12, "RANGE", "COUNT", "PROB", "CUM"
    Dim e(0 To 8) As Double
    For k = 0 To 8: e(k) = muF + (k - 4) * 0.75 * sdW: Next k
    Dim bc(0 To 9) As Long, bk As Long
    For i = 0 To m - 1
        bk = 9
        For k = 0 To 8
            If rr(i) <= e(k) Then bk = k: Exit For
        Next k
        bc(bk) = bc(bk) + 1
    Next i
    Dim lastBk As Long: lastBk = 9
    For k = 0 To 8
        If lastR <= e(k) Then lastBk = k: Exit For
    Next k
    Dim maxC As Long, cumC As Long
    For k = 0 To 9: If bc(k) > maxC Then maxC = bc(k)
    Next k
    For k = 0 To 9
        r = 33 + k
        cumC = cumC + bc(k)
        Dim lbl As String
        If k = 0 Then
            lbl = "<= " & Format(e(0) * 100, "0.00") & "%"
        ElseIf k = 9 Then
            lbl = "> " & Format(e(8) * 100, "0.00") & "%"
        Else
            lbl = Format(e(k - 1) * 100, "0.00") & " to " & Format(e(k) * 100, "0.00")
        End If
        PutC ws, r, 12, lbl, "@", VT_TXT, False, xlLeft
        PutC ws, r, 13, bc(k), "0", VT_TXT
        PutC ws, r, 14, bc(k) / m, "0.0%", VT_TXT
        PutC ws, r, 15, cumC / m, "0%", VT_DIM
        Dim midv As Double
        If k = 0 Then
            midv = e(0) - 1
        ElseIf k = 9 Then
            midv = e(8) + 1
        Else
            midv = (e(k - 1) + e(k)) / 2
        End If
        With ws.Cells(r, 16)
            .Value = String(CLng(14 * bc(k) / IIf(maxC > 0, maxC, 1)) + 1, ChrW(&H2588))
            .Font.Color = IIf(midv >= 0, VT_UP, VT_DOWN): .Font.Size = 7: .HorizontalAlignment = xlLeft
        End With
        If k = lastBk Then
            ws.Range(ws.Cells(r, 12), ws.Cells(r, 19)).Interior.Color = RGB(50, 34, 10)
            ws.Cells(r, 12).Font.Color = VT_OR: ws.Cells(r, 12).Font.Bold = True
        End If
    Next k

    ' --- biggest days ---
    Ttl ws, 44, 12, "BIGGEST DAYS", m & " " & unit, 14
    Hdr ws, 45, 12, "DATE", "WORST", "DATE", "BEST"
    Dim usedW() As Boolean, usedB() As Boolean
    ReDim usedW(0 To m - 1): ReDim usedB(0 To m - 1)
    For k = 0 To 4
        Dim bi As Long, wi As Long, bv As Double, wv As Double
        bv = -1E+99: wv = 1E+99: bi = -1: wi = -1
        For i = 0 To m - 1
            If Not usedB(i) And rr(i) > bv Then bv = rr(i): bi = i
            If Not usedW(i) And rr(i) < wv Then wv = rr(i): wi = i
        Next i
        r = 46 + k
        If wi >= 0 Then
            usedW(wi) = True
            PutC ws, r, 12, Format(CDate(dt(wi)), "yyyy-mm-dd"), "@", VT_TXT, False, xlLeft
            PutC ws, r, 13, wv, "+0.00%;-0.00%", ColOf(wv)
        End If
        If bi >= 0 Then
            usedB(bi) = True
            PutC ws, r, 14, Format(CDate(dt(bi)), "yyyy-mm-dd"), "@", VT_TXT, False, xlRight
            PutC ws, r, 15, bv, "+0.00%;-0.00%", ColOf(bv)
        End If
    Next k

    ' --- constituents (segment mode) ---
    If mode = "S" And gHasMem Then
        Ttl ws, 52, 12, "CONSTITUENTS", "last bar vs own daily SD  .  " & mCnt & " names", 14
        Hdr ws, 53, 12, "TICKER", "WEIGHT", "LAST BAR", "SD POS"
        Dim ord() As Long, t1 As Long
        ReDim ord(0 To mCnt - 1)
        For i = 0 To mCnt - 1: ord(i) = i: Next i
        For i = 1 To mCnt - 1
            t1 = ord(i): j = i - 1
            Do While j >= 0
                If mW(ord(j)) >= mW(t1) Then Exit Do
                ord(j + 1) = ord(j): j = j - 1
            Loop
            ord(j + 1) = t1
        Next i
        For i = 0 To mCnt - 1
            r = 54 + i
            k = ord(i)
            PutC ws, r, 12, mTk(k), "@", VT_TXT, True, xlLeft
            PutC ws, r, 13, mW(k), "0.0%", VT_DIM
            PutC ws, r, 14, mLast(k), "+0.00%;-0.00%", ColOf(mLast(k))
            If mSd(k) > 0 Then
                Dim zz As Double: zz = mLast(k) / mSd(k)
                PutC ws, r, 15, zz, "+0.0"" SD"";-0.0"" SD""", IIf(Abs(zz) >= 2, VT_OR, ColOf(zz)), Abs(zz) >= 2
            End If
        Next i
    End If

    gReady = True
End Sub

Private Function BinOf(ByVal x As Double, ByVal bw As Double) As Long
    Dim j As Long
    j = Int(x / bw + 0.5) + 20
    If j < 0 Then j = 0
    If j > NBIN - 1 Then j = NBIN - 1
    BinOf = j
End Function

' correlation / 60D correlation / beta / same direction against a date->return dictionary
Private Function ProxyStats(ByRef dd() As Long, ByRef rr() As Double, ByVal n As Long, ByVal dict As Object, _
                            ByRef corr As Double, ByRef corr60 As Double, ByRef beta As Double, _
                            ByRef same As Double, ByRef nA As Long) As Boolean
    Dim x() As Double, y() As Double, i As Long, m As Long
    If dict Is Nothing Then Exit Function
    ReDim x(0 To n): ReDim y(0 To n)
    For i = 0 To n - 1
        If dict.Exists(dd(i)) Then x(m) = rr(i): y(m) = dict(dd(i)): m = m + 1
    Next i
    If m < 20 Then Exit Function
    nA = m
    corr = CorrOf(x, y, 0, m - 1)
    Dim s0 As Long: s0 = m - 60: If s0 < 0 Then s0 = 0
    corr60 = CorrOf(x, y, s0, m - 1)
    Dim vy As Double: vy = SdS(y, 0, m - 1)
    If vy > 0 Then beta = corr * SdS(x, 0, m - 1) / vy
    Dim eq As Long
    For i = 0 To m - 1
        If (x(i) > 0 And y(i) > 0) Or (x(i) < 0 And y(i) < 0) Or (x(i) = 0 And y(i) = 0) Then eq = eq + 1
    Next i
    same = eq / m
    ProxyStats = True
End Function

Private Function CorrOf(ByRef x() As Double, ByRef y() As Double, ByVal i0 As Long, ByVal i1 As Long) As Double
    Dim mx As Double, my As Double, sxy As Double, sxx As Double, syy As Double, i As Long
    mx = MeanS(x, i0, i1): my = MeanS(y, i0, i1)
    For i = i0 To i1
        sxy = sxy + (x(i) - mx) * (y(i) - my)
        sxx = sxx + (x(i) - mx) ^ 2
        syy = syy + (y(i) - my) ^ 2
    Next i
    If sxx > 0 And syy > 0 Then CorrOf = sxy / Sqr(sxx * syy)
End Function

' segment dropdown list: names written far right, a workbook name points at them
Private Sub BuildSegList(ws As Worksheet)
    Dim a As Variant: a = SegNames()
    If IsEmpty(a) Then Exit Sub
    Dim i As Long, n As Long: n = UBound(a) - LBound(a) + 1
    Dim cc As Long: cc = DC + 16
    ws.Cells(1, cc).Value = "SEGMENTS"
    For i = 0 To n - 1
        ws.Cells(2 + i, cc).NumberFormat = "@"
        ws.Cells(2 + i, cc).Value = CStr(a(LBound(a) + i))
        ws.Cells(2 + i, cc).HorizontalAlignment = xlLeft
    Next i
    On Error Resume Next
    ThisWorkbook.Names("VtSegList").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="VtSegList", RefersTo:="='" & ws.Name & "'!" & _
        ws.Range(ws.Cells(2, cc), ws.Cells(1 + n, cc)).Address, Visible:=False
End Sub

'=============================================================================
'  Charts - called by the sheet code AFTER NavAdd (series references do not
'  follow a row / column insert, so they are built against the final layout)
'=============================================================================
Public Sub VtFinish(ws As Worksheet)
    If Not gReady Then Exit Sub
    Dim offR As Long, offC As Long
    offR = NavOffset(ws): offC = NavLeft(ws)
    Dim i As Long
    On Error Resume Next
    For i = ws.ChartObjects().count To 1 Step -1: ws.ChartObjects(i).Delete: Next i
    On Error GoTo 0

    Dim L As Double, T As Double, W As Double
    L = ws.Cells(1 + offR, 1 + offC).Left
    W = ws.Cells(1 + offR, 11 + offC).Left - L - 8
    T = ws.Cells(11 + offR, 1 + offC).Top

    ' ---- distribution ----
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(L, T, W, 330)
    co.Name = "VT_HIST": co.Placement = xlMove
    Dim c0 As Long: c0 = DC + offC
    With co.Chart
        Do While .SeriesCollection.count > 0: .SeriesCollection(1).Delete: Loop
        .ChartType = xlColumnClustered
        Dim s As Series
        Set s = .SeriesCollection.NewSeries
        s.Name = "COUNT"
        s.XValues = ws.Range(ws.Cells(2 + offR, c0), ws.Cells(1 + NBIN + offR, c0))
        s.Values = ws.Range(ws.Cells(2 + offR, c0 + 1), ws.Cells(1 + NBIN + offR, c0 + 1))
        s.Format.Fill.ForeColor.RGB = VT_DOWN
        s.Format.Line.Visible = msoFalse
        For i = 1 To NBIN
            If i - 1 > 20 Then
                s.Points(i).Format.Fill.ForeColor.RGB = VT_UP
            ElseIf i - 1 = 20 Then
                s.Points(i).Format.Fill.ForeColor.RGB = RGB(160, 160, 160)
            End If
        Next i
        Set s = .SeriesCollection.NewSeries
        s.Name = "LAST"
        s.XValues = ws.Range(ws.Cells(2 + offR, c0), ws.Cells(1 + NBIN + offR, c0))
        s.Values = ws.Range(ws.Cells(2 + offR, c0 + 3), ws.Cells(1 + NBIN + offR, c0 + 3))
        s.Format.Fill.ForeColor.RGB = VT_OR
        s.Format.Line.Visible = msoFalse
        Set s = .SeriesCollection.NewSeries
        s.Name = "NORMAL"
        s.XValues = ws.Range(ws.Cells(2 + offR, c0), ws.Cells(1 + NBIN + offR, c0))
        s.Values = ws.Range(ws.Cells(2 + offR, c0 + 2), ws.Cells(1 + NBIN + offR, c0 + 2))
        s.ChartType = xlLine
        s.Format.Line.ForeColor.RGB = RGB(235, 235, 235)
        s.Format.Line.Weight = 1.25
        s.Format.Line.DashStyle = msoLineDash
        s.MarkerStyle = xlMarkerStyleNone
        s.Smooth = True
        .ChartGroups(1).Overlap = 100
        .ChartGroups(1).GapWidth = 15
        Call StyleChart(co.Chart, gHistTitle, True)
        .Axes(xlCategory).TickLabelSpacing = 4
        .Axes(xlValue).TickLabels.NumberFormat = "0"
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .Legend.Font.Color = VT_DIM: .Legend.Font.Size = 8
    End With

    ' ---- window returns ----
    Dim T2 As Double: T2 = T + 338
    Set co = ws.ChartObjects.Add(L, T2, W, 280)
    co.Name = "VT_DAILY": co.Placement = xlMove
    Dim n As Long: n = gW
    With co.Chart
        Do While .SeriesCollection.count > 0: .SeriesCollection(1).Delete: Loop
        .ChartType = xlColumnClustered
        Dim nm As Variant, cl As Variant, k As Long
        nm = Array("UP", "DOWN", "BEYOND 2SD")
        cl = Array(VT_UP, VT_DOWN, VT_OR)
        For k = 0 To 2
            Set s = .SeriesCollection.NewSeries
            s.Name = nm(k)
            s.XValues = ws.Range(ws.Cells(2 + offR, c0 + 6), ws.Cells(1 + n + offR, c0 + 6))
            s.Values = ws.Range(ws.Cells(2 + offR, c0 + 7 + k), ws.Cells(1 + n + offR, c0 + 7 + k))
            s.Format.Fill.ForeColor.RGB = cl(k)
            s.Format.Line.Visible = msoFalse
        Next k
        nm = Array("+1SD", "+2SD", "-1SD", "-2SD")
        For k = 0 To 3
            Set s = .SeriesCollection.NewSeries
            s.Name = nm(k)
            s.XValues = ws.Range(ws.Cells(2 + offR, c0 + 6), ws.Cells(1 + n + offR, c0 + 6))
            s.Values = ws.Range(ws.Cells(2 + offR, c0 + 10 + k), ws.Cells(1 + n + offR, c0 + 10 + k))
            s.ChartType = xlLine
            s.Format.Line.ForeColor.RGB = IIf(k Mod 2 = 0, RGB(120, 120, 120), RGB(200, 200, 200))
            s.Format.Line.Weight = 0.75
            s.Format.Line.DashStyle = msoLineDash
            s.MarkerStyle = xlMarkerStyleNone
        Next k
        .ChartGroups(1).Overlap = 100
        .ChartGroups(1).GapWidth = 30
        Call StyleChart(co.Chart, gDailyTitle, False)
        Dim sp As Long: sp = n \ 8: If sp < 1 Then sp = 1
        .Axes(xlCategory).TickLabelSpacing = sp
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .HasLegend = False
    End With
End Sub

Private Sub StyleChart(ch As Chart, ByVal title As String, ByVal withLegend As Boolean)
    With ch
        .HasTitle = True
        With .ChartTitle
            .Text = title
            .Font.Size = 10: .Font.Bold = True: .Font.Color = RR4_ACCENT
            .Format.TextFrame2.TextRange.Font.Size = 10
        End With
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .ChartArea.Format.Line.ForeColor.RGB = RR4_LINE
        .ChartArea.Font.Name = "Consolas"
        .ChartArea.Font.Color = VT_TXT
        .PlotArea.Format.Fill.Visible = msoFalse
        With .Axes(xlCategory)
            .TickLabels.Font.Size = 7
            .TickLabels.Font.Color = VT_DIM
            .Format.Line.ForeColor.RGB = RR4_LINE
        End With
        With .Axes(xlValue)
            .TickLabels.Font.Size = 8
            .TickLabels.Font.Color = VT_DIM
            .Format.Line.ForeColor.RGB = RR4_LINE
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(35, 35, 35)
        End With
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
    End With
End Sub

'=============================================================================
'  Price history from Yahoo. A bare TW number is tried as .TW first and then
'  as .TWO - the old version only ever tried .TW, so OTC names came back empty.
'=============================================================================
Function GetHistoricalData(Ticker As String, ByRef outDates() As Date, ByRef outPrices() As Double) As Boolean
    Dim t As String
    t = UCase$(Trim$(Ticker))

    If InStr(t, ".") = 0 And IsNumeric(t) Then
        If FetchYahooSeries(t & ".TW", outDates, outPrices) Then
            GetHistoricalData = True
            Exit Function
        End If
        GetHistoricalData = FetchYahooSeries(t & ".TWO", outDates, outPrices)
    Else
        GetHistoricalData = FetchYahooSeries(t, outDates, outPrices)
    End If
End Function

Private Function FetchYahooSeries(ByVal Ticker As String, ByRef outDates() As Date, _
                                  ByRef outPrices() As Double) As Boolean
    Dim http As Object
    Dim url As String, response As String
    Dim tsStr As String, closeStr As String
    Dim tsArr() As String, closeArr() As String
    Dim i As Long

    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?range=10y&interval=1d"

    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        response = .responseText
    End With
    On Error GoTo 0
    If Len(response) = 0 Then Exit Function

    Dim tsStart As Long, tsEnd As Long
    tsStart = InStr(response, """timestamp"":[") + 12
    tsEnd = InStr(tsStart, response, "]")
    If tsStart <= 12 Or tsEnd <= tsStart Then Exit Function
    tsStr = Mid(response, tsStart, tsEnd - tsStart)
    tsArr = Split(tsStr, ",")

    Dim closeStart As Long, closeEnd As Long
    closeStart = InStr(response, """adjclose"":[{""adjclose"":[") + 23
    If closeStart < 24 Then closeStart = InStr(response, """close"":[") + 8
    closeEnd = InStr(closeStart, response, "]")
    If closeStart <= 20 Or closeEnd <= closeStart Then Exit Function
    closeStr = Mid(response, closeStart, closeEnd - closeStart)
    closeArr = Split(closeStr, ",")

    Dim count As Long
    count = UBound(tsArr)
    If UBound(closeArr) < count Then count = UBound(closeArr)
    If count < 1 Then Exit Function

    ReDim outDates(0 To count)
    ReDim outPrices(0 To count)

    Dim validCount As Long
    validCount = 0
    For i = 0 To count
        If IsNumeric(tsArr(i)) And IsNumeric(closeArr(i)) Then
            outDates(validCount) = (CDbl(tsArr(i)) / 86400) + 25569
            outPrices(validCount) = CDbl(closeArr(i))
            validCount = validCount + 1
        End If
    Next i

    If validCount > 1 Then
        ReDim Preserve outDates(0 To validCount - 1)
        ReDim Preserve outPrices(0 To validCount - 1)
        FetchYahooSeries = True
    End If
End Function
