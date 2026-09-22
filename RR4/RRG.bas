Attribute VB_Name = "RRG"
Option Explicit

' ================================================================
'  RELATIVE ROTATION GRAPH (nav code RG, recalc RG!) - sheet "RRG"
' ----------------------------------------------------------------
'  Prototype (2026-09-12): an Excel port of sector-rotation-system/
'  rrg_dynamic.py, snapshot only (no time slider).  The maths is kept
'  identical to the Python so the two can be checked against each other:
'
'    rs        = adjclose(ticker) / adjclose(SPY)           (daily)
'    rs_ratio  = 100 * rs / SMA(rs, 65)
'    rs_mom    = 100 * rs_ratio / SMA(rs_ratio, 20)         (unsmoothed ratio)
'    both      = EWM(span 3, adjust=True)  (pandas default)
'    tail      = last 66 rows, sampled every 5th trading day backwards
'                from the latest one -> 14 points (13 weeks)
'    quadrant  = ratio >= 100 ? (mom >= 100 ? LEADING : WEAKENING)
'                             : (mom >= 100 ? IMPROVING : LAGGING)
'
'  Universe = CorelationMatrix.SectorList() (config.py universe + SPY).
'  Prices come straight from the Yahoo chart API (1y OHLCV, high/low/close
'  scaled by adjclose/close the way yfinance auto_adjust does) so the
'  workbook does not depend on the Python side at all.
'
'  MONEY FLOW (money_flow_chart.py / money_flow_ranking.py, same page):
'    price chg = close[-1] / close[-21] - 1                (20 trading days)
'    CMF       = sum20(((C-L)-(H-C))/(H-L) * V) / sum20(V) (Chaikin, -1..1)
'    OBV trend = (OBV[-1] - OBV[-21]) / mean20(V)          (days of volume)
'    signal    = price dir x money dir (CMF dead zone 0.05, OBV 0.5 days;
'                CMF and OBV disagreeing -> money dir 0 + note)
'    INFLOW (up/in) / OUTFLOW (down/out) / BEAR DIV (up/out) /
'    BULL DIV (down/in) / NEUTRAL.  Marker: triangle OBV up, diamond OBV
'    down, circle flat.  Columns J:M of the table, chart RRG_FLOW.
'
'  Layout (page rows; the nav bar adds 3 on top - see NavOffset):
'    row 1   title + subtitle
'    row 2   hint
'    row 4   table header, rows 5..30 one ETF each (A..N; I = TRAIL sparkline)
'    P4      chart RRG_MAIN (scatter, one series per ETF)
'    P32     chart RRG_FLOW (money flow: price change vs CMF, one series per ETF)
'    row 33  quadrant + money-flow signal membership, then notes
'    Z5..    tail dates (shared by every ETF - all are aligned to SPY days)
'    AA4..   tail data block (2 columns per ETF, 14 rows) - chart source
'    after it: TRAIL block, one row per ETF, 13 weekly RS-RATIO changes (sparkline source)
'
'  FOCUS: double-click a ticker in column A -> only that trail stays lit
'  (thicker line, every dot dated), every other ETF goes dark grey and
'  loses its label, the table row gets the orange background.  Double-
'  click the same ticker again, or the TICKER header, to show all.  The
'  sheet's Worksheet_BeforeDoubleClick (RR4/SheetRRG_Code.txt) is written
'  into the sheet module by BuildRRG itself the first time the sheet is
'  built (needs "Trust access to the VBA project object model").  The
'  focused tickers live in the hidden sheet name RRGFOCUS (comma-separated:
'  double-click adds / removes; one ticker gets dated dots, several only names).
'  SORT: double-click a column header to sort the table by it (again to
'  flip); double-click the page title to restore the build order and clear
'  the focus.  See RrgSort.
'
'  INDUSTRY RRG (nav code RI, recalc RI!) - sheet "RRG Industry" (2026-09-13)
'    Same page, same maths, same focus / sort, but the universe is the 145
'    moomoo US industry plates and each "ticker" is a cap-weighted composite
'    of EVERY constituent (projects/moomoo-plate-list/moomoo_us_plate_stocks.csv,
'    5,812 stocks, imported into the IndustryMap sheet by ImportIndustryMap /
'    nav code IMAP).  Column A holds the plate code (BK2072), B the plate name;
'    chart labels show the name.  The composite is chain-linked from daily
'    returns so a stock that listed mid-year just joins the index when its
'    data starts:
'      idx[t] = idx[t-1] * (1 + sum_i w_i * r_i[t] / sum_i w_i)   over the
'               constituents that have both t-1 and t; w_i = market cap
'      high / low use the same weights on high[t]/close[t-1] - 1
'      volume = sum_i close_i * volume_i  (dollar volume, so CMF / OBV
'               are not adding up share counts of different stocks)
'    One RI! is ~5,800 Yahoo requests (40-60 min).  Each finished composite
'    is written to the hidden IndustryPx sheet keyed by plate code + as-of
'    day, so Esc (or a network drop) loses at most one industry: run RI!
'    again the same day and it continues from the cache.
'
'  TW GROUPS RRG (nav code TG, recalc TG!) - sheet "RRG TW Groups" (2026-09-13)
'    Third universe: the Market = TW rows of the Groups sheet (tblGroups,
'    the same thematic groups the Company research scanner uses - CCL,
'    CoWoS, MOSFET, ...), one cap-weighted composite per group vs ^TWII.
'    tblGroups has no market cap, so weights are shares outstanding x last
'    close: shares come from TWSE openapi t187ap03_L (listed) and the TPEx
'    daily close table (OTC - its openapi company table truncates at random,
'    do not use it), see TwSharesMap.  Tickers may be bare codes; the
'    exchange suffix is decided by which shares table lists the code, with
'    a .TW -> .TWO retry on a 404 as a fallback.  Everything else (maths,
'    composite, cache, focus, sort) is the industry page's.
'    Size floor (2026-09-22): a member whose market cap (shares x last
'    close) is below TW_MIN_CAP (NT$10bn) gets no weight and its volume is
'    left out of CMF / OBV too; a group with no member above the floor is
'    dropped from the page.  STOCKS = members counted / total, column B
'    lists the largest counted members.  The floor is part of the cache
'    key, so changing it refetches instead of reading stale composites.
' ================================================================

Private Const RS_WINDOW As Long = 65
Private Const MOM_WINDOW As Long = 20
Private Const EWM_SPAN As Long = 3
Private Const TAIL_WEEKS As Long = 13
Private Const TAIL_SPACING As Long = 5
Private Const TAIL_POINTS As Long = TAIL_WEEKS + 1
Private Const BENCH As String = "SPY"

Private Const TBL_HDR As Long = 4
Private Const TBL_FIRST As Long = 5
Private Const DATA_COL_MIN As Long = 27      ' AA: data block start on the ETF page (chart 560 wide fits before it)
Private Const DATA_COL_MARK As String = "RRGDATACOL"   ' hidden sheet name: where this page's data block starts
Private Const CHART_COL As Long = 16         ' P  (table is A:N, O is the gap)
Private Const TBL_NCOL As Long = 14
Private Const ETF_CHART_W As Double = 560
Private Const ETF_CHART_H As Double = 470
Private Const IND_CHART_W As Double = 1000
Private Const IND_CHART_H As Double = 820
Private gChartW As Double                    ' set per build (ETF / IND)
Private gChartH As Double

Private Const ETF_SHEET As String = "RRG"
Private Const IND_SHEET As String = "RRG Industry"
Private Const TWG_SHEET As String = "RRG TW Groups"
Private Const TWG_CHART_W As Double = 720
Private Const TWG_CHART_H As Double = 600
Private Const TW_BENCH As String = "^TWII"
Private Const TW_MIN_CAP As Double = 10000000000#   ' TW groups: members below NT$10bn market cap get no weight
Private gExclC As Long                       ' TW groups: members dropped by TW_MIN_CAP in the current build
Private gBench As String                     ' benchmark of the current build (SPY / ^TWII)
Private Const IND_MAP As String = "IndustryMap"      ' plateCode / plateName / symbol / rank / marketCap / plateType
Private Const IND_PX As String = "IndustryPx"        ' hidden cache: one composite series per industry row
Private Const PX_MAXN As Long = 270                  ' points per array in the cache (1y ~ 252 days)
Private Const PX_FIRST As Long = 6                   ' cache row: A code, B asof, C ok, D fail, E npts, F.. arrays

Private Const CMF_WINDOW As Long = 20
Private Const FLOW_LOOKBACK As Long = 20
Private Const CMF_DEADZONE As Double = 0.05
Private Const OBV_DEADZONE As Double = 0.5
Private Const FOCUS_MARK As String = "RRGFOCUS"
Private Const SORT_MARK As String = "RRGSORT"      ' "col|dir" of the current table sort
' SEQ (build order) and DATE columns sit just left of the data block: DataCol - 2 / DataCol - 1
Private Const DIM_GREY As Long = 4210752     ' RGB(64,64,64)

Private Const NAN As Double = -1E+300

Private Function QuadColor(ByVal q As String) As Long
    Select Case q
        Case "LEADING":   QuadColor = RGB(46, 125, 50)
        Case "WEAKENING": QuadColor = RGB(249, 168, 37)
        Case "LAGGING":   QuadColor = RGB(198, 40, 40)
        Case Else:        QuadColor = RGB(21, 101, 192)     ' IMPROVING
    End Select
End Function

' money_flow_ranking.SIGNAL_COLORS: red inflow / green outflow (TW
' convention, same as the weight bars), amber / blue for the divergences
Private Function SigColor(ByVal sig As String) As Long
    Select Case sig
        Case "INFLOW":   SigColor = RGB(214, 69, 69)
        Case "OUTFLOW":  SigColor = RGB(30, 142, 90)
        Case "BEAR DIV": SigColor = RGB(249, 168, 37)
        Case "BULL DIV": SigColor = RGB(21, 101, 192)
        Case Else:       SigColor = RGB(117, 117, 117)     ' NEUTRAL / N/A
    End Select
End Function

' money_flow_ranking.classify()
Private Function FlowSignal(ByVal pchg As Double, ByVal cmf As Double, ByVal obvT As Double, ByRef note As String) As String
    Dim pd As Long, cd As Long, od As Long, md As Long
    pd = IIf(pchg > 0, 1, IIf(pchg < 0, -1, 0))
    cd = IIf(cmf > CMF_DEADZONE, 1, IIf(cmf < -CMF_DEADZONE, -1, 0))
    od = IIf(obvT > OBV_DEADZONE, 1, IIf(obvT < -OBV_DEADZONE, -1, 0))
    note = ""
    If cd <> 0 And od <> 0 And cd <> od Then
        md = 0: note = "CMF/OBV disagree"
    Else
        md = IIf(cd <> 0, cd, od)
    End If
    If pd > 0 And md > 0 Then
        FlowSignal = "INFLOW"
    ElseIf pd < 0 And md < 0 Then
        FlowSignal = "OUTFLOW"
    ElseIf pd > 0 And md < 0 Then
        FlowSignal = "BEAR DIV"
    ElseIf pd < 0 And md > 0 Then
        FlowSignal = "BULL DIV"
    Else
        FlowSignal = "NEUTRAL"
    End If
End Function

Private Function Quadrant(ByVal x As Double, ByVal y As Double) As String
    If x >= 100 Then
        Quadrant = IIf(y >= 100, "LEADING", "WEAKENING")
    Else
        Quadrant = IIf(y >= 100, "IMPROVING", "LAGGING")
    End If
End Function

' ----------------------------------------------------------------
' One numeric JSON array out of the Yahoo chart response: the text between
' "<key>":[ and the next ], split on commas (nulls come back as "null").
Private Function JsonArr(ByRef resp As String, ByVal key As String, ByVal fromPos As Long, ByRef parts() As String) As Boolean
    Dim p1 As Long, p2 As Long
    p1 = InStr(fromPos, resp, """" & key & """:[")
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(key) + 4
    p2 = InStr(p1, resp, "]")
    If p2 = 0 Then Exit Function
    parts = Split(Mid(resp, p1, p2 - p1), ",")
    JsonArr = True
End Function

' Yahoo chart API, 1y daily OHLCV.  c() is adjclose; h() / l() are the raw
' high / low scaled by adjclose / close (what yfinance auto_adjust does);
' v() is the raw volume.  Days where any field is null are dropped, so the
' arrays stay aligned with each other.  Returns the number of rows.
Private Function FetchOhlcv(ByVal Ticker As String, ByRef days() As Long, ByRef c() As Double, _
                            ByRef h() As Double, ByRef l() As Double, ByRef v() As Double) As Long
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Replace(Ticker, "^", "%5E") & "?interval=1d&range=1y"
    ' a few thousand requests in a row (the industry page) do hit 429s /
    ' dropped connections now and then: back off and retry before giving up
    Dim attempt As Long, status As Long
    For attempt = 1 To 3
        status = 0
        On Error Resume Next
        http.Open "GET", url, False
        http.setRequestHeader "User-Agent", "Mozilla/5.0"
        http.send
        status = http.Status
        On Error GoTo 0
        If status = 200 Then Exit For
        If status = 404 Then Exit Function                 ' unknown symbol: no point retrying
        Application.Wait Now + TimeSerial(0, 0, 2 * attempt)
        Set http = CreateObject("MSXML2.XMLHTTP")
    Next attempt
    If status <> 200 Then Exit Function
    Dim resp As String: resp = http.responseText

    Dim ts() As String, cl() As String, hi() As String, lo() As String, vo() As String, ac() As String
    If Not JsonArr(resp, "timestamp", 1, ts) Then Exit Function
    If Not JsonArr(resp, "close", 1, cl) Then Exit Function          ' "adjclose":[ does not match "close":[
    If Not JsonArr(resp, "high", 1, hi) Then Exit Function
    If Not JsonArr(resp, "low", 1, lo) Then Exit Function
    If Not JsonArr(resp, "volume", 1, vo) Then Exit Function
    ' the numeric adjclose array is the SECOND "adjclose":[ (the first opens the object)
    Dim pAdj As Long: pAdj = InStr(resp, """adjclose"":[")
    If pAdj = 0 Then Exit Function
    If Not JsonArr(resp, "adjclose", pAdj + 1, ac) Then Exit Function

    Dim n As Long: n = UBound(ts) + 1
    If UBound(cl) + 1 < n Then n = UBound(cl) + 1
    If UBound(hi) + 1 < n Then n = UBound(hi) + 1
    If UBound(lo) + 1 < n Then n = UBound(lo) + 1
    If UBound(vo) + 1 < n Then n = UBound(vo) + 1
    If UBound(ac) + 1 < n Then n = UBound(ac) + 1
    If n <= 0 Then Exit Function
    ReDim days(0 To n - 1): ReDim c(0 To n - 1): ReDim h(0 To n - 1): ReDim l(0 To n - 1): ReDim v(0 To n - 1)
    Dim k As Long, cnt As Long
    For k = 0 To n - 1
        If IsNumeric(Trim(ts(k))) And IsNumeric(Trim(cl(k))) And IsNumeric(Trim(hi(k))) And _
           IsNumeric(Trim(lo(k))) And IsNumeric(Trim(vo(k))) And IsNumeric(Trim(ac(k))) Then
            Dim rawC As Double: rawC = CDbl(Val(cl(k)))
            If rawC > 0 Then
                Dim ratio As Double: ratio = CDbl(Val(ac(k))) / rawC
                days(cnt) = CLng(Val(ts(k)) \ 86400)
                c(cnt) = CDbl(Val(ac(k)))
                h(cnt) = CDbl(Val(hi(k))) * ratio
                l(cnt) = CDbl(Val(lo(k))) * ratio
                v(cnt) = CDbl(Val(vo(k)))
                cnt = cnt + 1
            End If
        End If
    Next k
    FetchOhlcv = cnt
End Function

' ----------------------------------------------------------------
Sub BuildRRG()
    Call BuildRRGCore("ETF")
End Sub

' RI!: the industry page.  limitN > 0 builds only the first limitN
' industries (testing - a full run is ~5,800 requests).
Sub BuildRRGIndustry(Optional ByVal limitN As Long = 0)
    Call BuildRRGCore("IND", limitN)
End Sub

' TG!: the Taiwan groups page (tblGroups, Market = TW).
Sub BuildRRGTwGroups(Optional ByVal limitN As Long = 0)
    Call BuildRRGCore("TWG", limitN)
End Sub

' kind = "ETF" (sheet RRG, SectorList universe, one Yahoo series per ETF)
'     or "IND" (sheet RRG Industry, IndustryMap universe, composites)
Private Sub BuildRRGCore(ByVal kind As String, Optional ByVal limitN As Long = 0)
    Dim isInd As Boolean: isInd = (kind = "IND")
    Dim isTwg As Boolean: isTwg = (kind = "TWG")
    Dim isComp As Boolean: isComp = isInd Or isTwg           ' composite universes
    Dim sheetNm As String, navCode As String
    Select Case kind
        Case "IND": sheetNm = IND_SHEET: navCode = "RGI": gChartW = IND_CHART_W: gChartH = IND_CHART_H: gBench = BENCH
        Case "TWG": sheetNm = TWG_SHEET: navCode = "RGT": gChartW = TWG_CHART_W: gChartH = TWG_CHART_H: gBench = TW_BENCH
        Case Else:  sheetNm = ETF_SHEET: navCode = "RGE": gChartW = ETF_CHART_W: gChartH = ETF_CHART_H: gBench = BENCH
    End Select
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetNm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = sheetNm
    End If
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Dim prevCancel As Long: prevCancel = Application.EnableCancelKey
    Application.EnableCancelKey = xlErrorHandler        ' Esc -> error 18, handled below (cache keeps the progress)
    On Error GoTo Fail
    Call NavStrip(ws)

    ' ---- universe ----
    Dim tickers() As String, labels() As String, groups() As String
    Dim i As Long, j As Long, k As Long, n As Long
    Dim indSyms As Variant, indCaps As Variant          ' IND: per industry, its constituents + weights
    Dim nStocks As Long
    If isInd Then
        n = IndustryUniverse(tickers, labels, groups, indSyms, indCaps, nStocks)
        If n = 0 Then Err.Raise vbObjectError + 2, , IND_MAP & " sheet is empty - run IMAP (ImportIndustryMap) first"
        If limitN > 0 And limitN < n Then n = limitN
    ElseIf isTwg Then
        n = TwGroupUniverse(tickers, labels, groups, indSyms, indCaps, nStocks)
        If n = 0 Then Err.Raise vbObjectError + 3, , "Groups sheet has no Market = TW rows"
        If limitN > 0 And limitN < n Then n = limitN
    Else
        ' SPY last in the list is the benchmark
        Dim lst As Variant: lst = SectorList()
        Dim nAll As Long: nAll = UBound(lst) + 1
        n = 0
        ReDim tickers(0 To nAll - 1): ReDim labels(0 To nAll - 1): ReDim groups(0 To nAll - 1)
        For i = 0 To nAll - 1
            If UCase(lst(i)(0)) <> BENCH Then
                tickers(n) = lst(i)(0): labels(n) = lst(i)(1): groups(n) = lst(i)(2)
                n = n + 1
            End If
        Next i
        ReDim Preserve tickers(0 To n - 1): ReDim Preserve labels(0 To n - 1): ReDim Preserve groups(0 To n - 1)
    End If
    Dim unitNm As String: unitNm = IIf(isInd, "industries", IIf(isTwg, "groups", "ETFs"))

    ' ---- benchmark ----
    Dim bDays() As Long, bPx() As Double, nb As Long
    Dim bH() As Double, bL() As Double, bV() As Double
    Application.StatusBar = "RRG: fetching " & gBench
    nb = FetchOhlcv(gBench, bDays, bPx, bH, bL, bV)
    If nb < RS_WINDOW + MOM_WINDOW + 5 Then Err.Raise vbObjectError + 1, , "not enough " & gBench & " data (" & nb & " days)"
    Dim bIdx As Object: Set bIdx = CreateObject("Scripting.Dictionary")
    For i = 0 To nb - 1: bIdx(bDays(i)) = i: Next i

    ' ---- per ticker: aligned rs -> ratio / momentum -> tail ----
    Dim tailX() As Double, tailY() As Double, tailN() As Long
    ReDim tailX(0 To n - 1, 0 To TAIL_POINTS - 1)
    ReDim tailY(0 To n - 1, 0 To TAIL_POINTS - 1)
    ReDim tailN(0 To n - 1)
    Dim rs() As Double, ratio() As Double, mom() As Double, sr() As Double, sm() As Double
    Dim vx() As Double, vy() As Double, nv As Long
    Dim tDays() As Long, tPx() As Double, nt As Long
    Dim tH() As Double, tL() As Double, tV() As Double
    ' money flow per ticker (on the ticker's own rows, like the python)
    Dim fPchg() As Double, fCmf() As Double, fObv() As Double, fSig() As String, fNote() As String, fOk() As Boolean
    ReDim fPchg(0 To n - 1): ReDim fCmf(0 To n - 1): ReDim fObv(0 To n - 1)
    ReDim fSig(0 To n - 1): ReDim fNote(0 To n - 1): ReDim fOk(0 To n - 1)
    Dim t0 As Double: t0 = Timer
    Dim okC As Long, failC As Long
    gExclC = 0
    Dim keep() As Boolean: ReDim keep(0 To n - 1)
    For i = 0 To n - 1
        Application.StatusBar = "RRG: fetching [" & (i + 1) & "/" & n & "] " & tickers(i)
        DoEvents
        If isComp Then
            Dim okThis As Long
            nt = IndustryComposite(tickers(i), labels(i), indSyms(i), indCaps(i), isTwg, bDays, nb, bIdx, i + 1, n, t0, _
                                   tDays, tPx, tH, tL, tV, okC, failC, okThis)
            groups(i) = okThis & "/" & (UBound(indSyms(i)) + 1)
            keep(i) = Not (isTwg And okThis = 0)         ' TW: no member above TW_MIN_CAP -> drop the group
        Else
            nt = FetchOhlcv(tickers(i), tDays, tPx, tH, tL, tV)
            keep(i) = True
        End If
        fOk(i) = MoneyFlow(tPx, tH, tL, tV, nt, fPchg(i), fCmf(i), fObv(i), fSig(i), fNote(i))
        ReDim rs(0 To nb - 1): ReDim ratio(0 To nb - 1): ReDim mom(0 To nb - 1)
        ReDim sr(0 To nb - 1): ReDim sm(0 To nb - 1)
        For k = 0 To nb - 1: rs(k) = NAN: ratio(k) = NAN: mom(k) = NAN: sr(k) = NAN: sm(k) = NAN: Next k
        For k = 0 To nt - 1
            If bIdx.Exists(tDays(k)) Then rs(bIdx(tDays(k))) = tPx(k) / bPx(bIdx(tDays(k)))
        Next k
        Call RollingRatio(rs, ratio, RS_WINDOW)
        Call RollingRatio(ratio, mom, MOM_WINDOW)
        Call EwmAdjust(ratio, sr, EWM_SPAN)
        Call EwmAdjust(mom, sm, EWM_SPAN)
        ' rows where both are valid (pandas dropna), keep the last 66
        ReDim vx(0 To nb - 1): ReDim vy(0 To nb - 1): nv = 0
        For k = 0 To nb - 1
            If sr(k) <> NAN And sm(k) <> NAN Then vx(nv) = sr(k): vy(nv) = sm(k): nv = nv + 1
        Next k
        Dim span As Long: span = TAIL_WEEKS * TAIL_SPACING + 1
        If nv < span Then span = nv
        ' sample backwards from the newest point every TAIL_SPACING rows
        Dim cnt As Long: cnt = 0
        For k = nv - 1 To nv - span Step -TAIL_SPACING
            cnt = cnt + 1
        Next k
        tailN(i) = cnt
        j = TAIL_POINTS - 1
        For k = nv - 1 To nv - span Step -TAIL_SPACING
            tailX(i, j) = vx(k): tailY(i, j) = vy(k): j = j - 1
        Next k
    Next i
    Application.StatusBar = False

    ' ---- drop groups with nothing left (TW size floor) ----
    Dim dropped As String, o As Long
    o = 0
    For i = 0 To n - 1
        If keep(i) Then
            If o <> i Then
                tickers(o) = tickers(i): labels(o) = labels(i): groups(o) = groups(i)
                tailN(o) = tailN(i)
                For k = 0 To TAIL_POINTS - 1: tailX(o, k) = tailX(i, k): tailY(o, k) = tailY(i, k): Next k
                fPchg(o) = fPchg(i): fCmf(o) = fCmf(i): fObv(o) = fObv(i)
                fSig(o) = fSig(i): fNote(o) = fNote(i): fOk(o) = fOk(i)
            End If
            o = o + 1
        Else
            dropped = dropped & IIf(dropped = "", "", ", ") & tickers(i)
        End If
    Next i
    Dim nDropped As Long: nDropped = n - o
    n = o
    If n = 0 Then Err.Raise vbObjectError + 4, , "no group has a member above the size floor"

    ' ---- page ----
    Application.ScreenUpdating = False
    Call ShapesMoveOnly(ws)
    Dim co As ChartObject
    For Each co In ws.ChartObjects: co.Delete: Next co
    ws.cells.SparklineGroups.Clear
    ws.cells.Clear
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"                 ' table stays Consolas on every page
        .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.DisplayGridlines = False
    Dim rr As Long
    For rr = 1 To TBL_FIRST + n + 40: ws.Rows(rr).RowHeight = 18: Next rr
    ws.Columns(1).ColumnWidth = IIf(isTwg, 22, 8): ws.Columns(2).ColumnWidth = IIf(isInd, 34, IIf(isTwg, 26, 9))
    ws.Columns(3).ColumnWidth = IIf(isComp, 8, 15)
    ws.Columns(4).ColumnWidth = 9: ws.Columns(5).ColumnWidth = 9: ws.Columns(6).ColumnWidth = 11
    ws.Columns(7).ColumnWidth = 9: ws.Columns(8).ColumnWidth = 9: ws.Columns(9).ColumnWidth = 10
    ws.Columns(10).ColumnWidth = 5: ws.Columns(11).ColumnWidth = 9: ws.Columns(12).ColumnWidth = 8
    ws.Columns(13).ColumnWidth = 8: ws.Columns(14).ColumnWidth = 10: ws.Columns(15).ColumnWidth = 3

    Dim asOf As Date: asOf = DateSerial(1970, 1, 1) + bDays(nb - 1)
    With ws.cells(1, 1)
        .Value = IIf(isInd, "RRG  INDUSTRIES", IIf(isTwg, "RRG  TW GROUPS", "RELATIVE ROTATION GRAPH"))
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 14
    End With
    ws.Rows(1).RowHeight = 24
    With ws.cells(1, 4)
        .Value = n & " " & unitNm & IIf(isComp, " (" & okC & " stocks, cap-weighted composites" & _
                 IIf(isTwg, ", " & gExclC & " below NT$" & Format(TW_MIN_CAP / 100000000#, "0") & "e" & _
                     IIf(nDropped > 0, ", " & nDropped & " groups dropped", ""), "") & _
                 IIf(failC > 0, ", " & failC & " no data", "") & ")", "") & _
                 " vs " & gBench & "  .  daily adjclose  .  RS " & RS_WINDOW & "d / MOM " & MOM_WINDOW & _
                 "d / EWM " & EWM_SPAN & "  .  " & TAIL_WEEKS & "-week tail  .  as of " & Format(asOf, "yyyy/mm/dd") & _
                 "  .  built " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RGB(120, 120, 120): .Font.Size = 9
    End With
    With ws.cells(2, 1)
        .Value = navCode & "! <GO> refetches and redraws" & IIf(isInd, " (~5,800 requests, 40-60 min; Esc stops, the same-day cache resumes)", "") & _
                 "  .  double-click " & IIf(isInd, "codes", IIf(isTwg, "groups", "tickers")) & " to focus them (again to remove, title to show all)  .  double-click a header to sort (again to flip)  .  double-click the title to reset"
        .Font.Color = RGB(120, 120, 120): .Font.Size = 9
    End With
    With ws.Range(ws.cells(2, 1), ws.cells(2, TBL_NCOL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With

    ' ---- table ----
    Dim hdr As Variant
    hdr = Array(IIf(isInd, "CODE", IIf(isTwg, "GROUP", "TICKER")), IIf(isInd, "INDUSTRY", IIf(isTwg, "MEMBERS", "LABEL")), IIf(isComp, "STOCKS", "GROUP"), _
                "RS-RATIO", "RS-MOM", "QUADRANT", "1W dRAT", "1W dMOM", "TRAIL", "PTS", _
                "20D PX%", "CMF", "OBV(d)", "SIGNAL")
    For j = 0 To UBound(hdr)
        With ws.cells(TBL_HDR, j + 1)
            .Value = hdr(j)
            .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 9
            .HorizontalAlignment = IIf(j >= 3, xlRight, xlLeft)
        End With
    Next j
    With ws.Range(ws.cells(TBL_HDR, 1), ws.cells(TBL_HDR, TBL_NCOL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With
    Dim quadList(0 To 3) As String
    quadList(0) = "LEADING": quadList(1) = "IMPROVING": quadList(2) = "WEAKENING": quadList(3) = "LAGGING"
    Dim quadMembers(0 To 3) As String
    Dim sigList(0 To 4) As String
    sigList(0) = "INFLOW": sigList(1) = "BULL DIV": sigList(2) = "BEAR DIV": sigList(3) = "OUTFLOW": sigList(4) = "NEUTRAL"
    Dim sigMembers(0 To 4) As String
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    xMin = 99: xMax = 101: yMin = 99: yMax = 101
    For i = 0 To n - 1
        Dim r As Long: r = TBL_FIRST + i
        ws.cells(r, 1).Value = tickers(i)
        ws.cells(r, 2).Value = labels(i)
        If isComp Then ws.cells(r, 3).NumberFormat = "@"      ' "8/8" would otherwise become a date
        ws.cells(r, 3).Value = groups(i)
        If tailN(i) = 0 Then
            ws.cells(r, 6).Value = "NO DATA"
        Else
            Dim lx As Double, ly As Double
            lx = tailX(i, TAIL_POINTS - 1): ly = tailY(i, TAIL_POINTS - 1)
            Dim q As String: q = Quadrant(lx, ly)
            ws.cells(r, 4).Value = lx: ws.cells(r, 5).Value = ly
            ws.Range(ws.cells(r, 4), ws.cells(r, 5)).NumberFormat = "0.00"
            ws.cells(r, 6).Value = q
            If tailN(i) >= 2 Then
                ws.cells(r, 7).Value = lx - tailX(i, TAIL_POINTS - 2)
                ws.cells(r, 8).Value = ly - tailY(i, TAIL_POINTS - 2)
                ws.Range(ws.cells(r, 7), ws.cells(r, 8)).NumberFormat = "+0.00;-0.00;0.00"
            End If
            ws.cells(r, 10).Value = tailN(i)
            For j = 0 To 3
                If quadList(j) = q Then quadMembers(j) = quadMembers(j) & IIf(quadMembers(j) = "", "", "  ") & tickers(i)
            Next j
            ' axis range over every tail point (python: min/max of all in-range rows + 1.5 pad)
            For k = TAIL_POINTS - tailN(i) To TAIL_POINTS - 1
                If tailX(i, k) < xMin Then xMin = tailX(i, k)
                If tailX(i, k) > xMax Then xMax = tailX(i, k)
                If tailY(i, k) < yMin Then yMin = tailY(i, k)
                If tailY(i, k) > yMax Then yMax = tailY(i, k)
            Next k
        End If
        ' money flow columns J:M
        If fOk(i) Then
            ws.cells(r, 11).Value = fPchg(i): ws.cells(r, 11).NumberFormat = "+0.0%;-0.0%;0.0%"
            ws.cells(r, 12).Value = fCmf(i): ws.cells(r, 12).NumberFormat = "+0.000;-0.000;0.000"
            ws.cells(r, 13).Value = fObv(i): ws.cells(r, 13).NumberFormat = "+0.0;-0.0;0.0"
            ws.cells(r, 14).Value = fSig(i)
            If fNote(i) <> "" Then ws.cells(r, 14).AddComment fNote(i)
            For j = 0 To 4
                If sigList(j) = fSig(i) Then sigMembers(j) = sigMembers(j) & IIf(sigMembers(j) = "", "", "  ") & tickers(i)
            Next j
        Else
            ws.cells(r, 14).Value = "N/A"
        End If
        ws.Range(ws.cells(r, 4), ws.cells(r, TBL_NCOL)).HorizontalAlignment = xlRight
        ws.cells(r, 6).HorizontalAlignment = xlRight
    Next i
    Dim lastRow As Long: lastRow = TBL_FIRST + n - 1
    With ws.Range(ws.cells(lastRow, 1), ws.cells(lastRow, TBL_NCOL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With
    Call PaintTableRows(ws, "")

    ' ---- quadrant membership ----
    Dim qr As Long: qr = lastRow + 2
    With ws.cells(qr, 1)
        .Value = "QUADRANTS": .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With
    For j = 0 To 3
        With ws.cells(qr + 1 + j, 1)
            .Value = quadList(j): .Font.Color = QuadColor(quadList(j)): .Font.Bold = True
        End With
        ws.cells(qr + 1 + j, 3).Value = IIf(quadMembers(j) = "", "-", quadMembers(j))
    Next j

    Dim fr As Long: fr = qr + 6
    With ws.cells(fr, 1)
        .Value = "MONEY FLOW SIGNALS": .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With
    For j = 0 To 4
        With ws.cells(fr + 1 + j, 1)
            .Value = sigList(j): .Font.Color = SigColor(sigList(j)): .Font.Bold = True
        End With
        ws.cells(fr + 1 + j, 3).Value = IIf(sigMembers(j) = "", "-", sigMembers(j))
    Next j

    ' ---- notes ----
    Dim nr As Long: nr = fr + 7
    With ws.cells(nr, 1)
        .Value = "HOW THESE ARE COMPUTED": .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With
    Dim notes As Variant
    notes = Array( _
        "RS-RATIO  = 100 x (px / " & gBench & ") / SMA65(px / " & gBench & "), then EWM span 3  -  relative strength vs its own quarter trend", _
        "RS-MOM    = 100 x RS-RATIO / SMA20(RS-RATIO), then EWM span 3  -  is that strength accelerating (>100) or fading", _
        "TAIL      = the last 13 weeks, one dot per 5 trading days; the big dot is today, older dots fade", _
        "QUADRANT  = LEADING (both >= 100) / WEAKENING (ratio >= 100, mom < 100) / LAGGING (both < 100) / IMPROVING (ratio < 100, mom >= 100)", _
        "1W dRAT / dMOM = change since the previous tail dot (5 trading days), the direction the tail is heading", _
        "Same windows and smoothing as sector-rotation-system/rrg_dynamic.py; prices are Yahoo adjclose, so dividends are in.", _
        "", _
        "MONEY FLOW (chart below the RRG; money_flow_chart.py / money_flow_ranking.py)", _
        "20D PX%   = close today / close 20 trading days ago - 1                      (x axis)", _
        "CMF       = sum20( ((C-L)-(H-C))/(H-L) x V ) / sum20(V)  -  Chaikin Money Flow, -1..+1  (y axis)", _
        "            where each day's close sits in its own high-low range, volume weighted: >0 net buying, <0 net selling", _
        "OBV(d)    = (OBV today - OBV 20 days ago) / avg volume, in days of typical volume; sign is what matters", _
        "            OBV adds the whole day's volume on an up close and subtracts it on a down close (volume leads price)", _
        "SIGNAL    = price direction x money direction; CMF within +/-0.05 and OBV within +/-0.5 days count as flat,", _
        "            CMF and OBV pointing opposite ways = money direction unknown (cell comment says CMF/OBV disagree)", _
        "            INFLOW up+in / OUTFLOW down+out / BEAR DIV up but money out / BULL DIV down but money in / NEUTRAL", _
        "MARKER    = triangle OBV rising, diamond OBV falling, circle flat; colour = SIGNAL; shaded band = CMF dead zone")
    If isInd Then
        Dim indNotes As Variant
        indNotes = Array( _
            "", _
            "INDUSTRY COMPOSITES (this page)", _
            "Universe  = the 145 moomoo US industry plates; every constituent is used (IndustryMap sheet, IMAP to reimport the CSV)", _
            "Price     = cap-weighted index chain-linked from daily returns: idx[t] = idx[t-1] x (1 + sum w x r / sum w) over the", _
            "            stocks that have both days; a stock that listed mid-year joins when its data starts.  Weights = market cap.", _
            "High/Low  = same weights on high[t]/close[t-1] - 1 and low[t]/close[t-1] - 1;  Volume = sum(close x volume) in dollars", _
            "Cache     = hidden IndustryPx sheet, one composite per industry keyed by as-of day; RGI! reuses it the same day (Esc resumes)", _
            "STOCKS    = constituents with data / total in the plate")
        Dim tmpN As Variant: tmpN = notes
        ReDim Preserve tmpN(0 To UBound(notes) + UBound(indNotes) + 1)
        For j = 0 To UBound(indNotes): tmpN(UBound(notes) + 1 + j) = indNotes(j): Next j
        notes = tmpN
    End If
    If isTwg Then
        Dim twNotes As Variant
        twNotes = Array( _
            "", _
            "TW GROUP COMPOSITES (this page)", _
            "Universe  = the Market = TW rows of the Groups sheet (tblGroups) - the Company research scanner's thematic groups", _
            "Price     = cap-weighted index chain-linked from daily returns (see RRG.bas); weight = shares outstanding x last close", _
            "            shares: TWSE openapi t187ap03_L for listed, TPEx daily close table for OTC; a code in neither gets weight 0", _
            "Size      = members below NT$" & Format(TW_MIN_CAP / 100000000#, "0") & "e market cap get no weight and no volume; a group with none above it is dropped", _
            "High/Low  = same weights on high[t]/close[t-1] - 1 and low[t]/close[t-1] - 1;  Volume = sum(close x volume) in TWD", _
            "Benchmark = " & gBench & " (Yahoo);  tickers try .TW then .TWO like the scanner", _
            "Cache     = hidden IndustryPx sheet keyed by group name + size floor + as-of day; RGT! reuses it the same day", _
            "STOCKS    = members counted (data and above the size floor) / total in the group;  B = largest counted members")
        Dim tmpT As Variant: tmpT = notes
        ReDim Preserve tmpT(0 To UBound(notes) + UBound(twNotes) + 1)
        For j = 0 To UBound(twNotes): tmpT(UBound(notes) + 1 + j) = twNotes(j): Next j
        notes = tmpT
    End If
    For j = 0 To UBound(notes)
        ws.cells(nr + 1 + j, 1).Value = notes(j)
        ws.cells(nr + 1 + j, 1).Font.Color = RGB(150, 150, 150)
        ws.cells(nr + 1 + j, 1).Font.Size = 8
    Next j

    ' ---- tail data block (chart source): starts past the charts' right edge ----
    Dim dcol As Long: dcol = DATA_COL_MIN
    Do While ws.Columns(dcol - 2).Left < ws.Columns(CHART_COL).Left + gChartW + 12
        dcol = dcol + 1
    Loop
    Call SetMark(ws, DATA_COL_MARK, CStr(dcol))
    With ws.cells(TBL_HDR - 1, dcol)
        .Value = "TAIL DATA (chart source, oldest -> newest)": .Font.Color = RGB(90, 90, 90): .Font.Size = 8
    End With
    ' tail dates: every ETF is aligned to SPY's days and sampled backwards
    ' from the same last day, so one date column serves all of them
    ws.cells(TBL_HDR, SeqCol(ws)).Value = "seq": ws.cells(TBL_HDR, SeqCol(ws)).Font.Color = RGB(90, 90, 90)
    For i = 0 To n - 1: ws.cells(TBL_FIRST + i, SeqCol(ws)).Value = i + 1: Next i
    ws.Range(ws.cells(TBL_FIRST, SeqCol(ws)), ws.cells(TBL_FIRST + n - 1, SeqCol(ws))).Font.Color = RGB(90, 90, 90)
    ws.Columns(SeqCol(ws)).ColumnWidth = 4
    ws.cells(TBL_HDR, DateCol(ws)).Value = "date": ws.cells(TBL_HDR, DateCol(ws)).Font.Color = RGB(90, 90, 90)
    For k = 0 To TAIL_POINTS - 1
        Dim di As Long: di = nb - 1 - TAIL_SPACING * (TAIL_POINTS - 1 - k)
        If di >= 0 Then ws.cells(TBL_FIRST + k, DateCol(ws)).Value = DateSerial(1970, 1, 1) + bDays(di)
    Next k
    ws.Range(ws.cells(TBL_FIRST, DateCol(ws)), ws.cells(TBL_FIRST + TAIL_POINTS - 1, DateCol(ws))).NumberFormat = "mm/dd"
    ws.Range(ws.cells(TBL_FIRST, DateCol(ws)), ws.cells(TBL_FIRST + TAIL_POINTS - 1, DateCol(ws))).Font.Color = RGB(90, 90, 90)
    ws.Columns(DateCol(ws)).ColumnWidth = 7
    For i = 0 To n - 1
        Dim cx As Long: cx = DataCol(ws) + 2 * i
        ws.cells(TBL_HDR, cx).Value = tickers(i): ws.cells(TBL_HDR, cx + 1).Value = "mom"
        ws.Range(ws.cells(TBL_HDR, cx), ws.cells(TBL_HDR, cx + 1)).Font.Color = RGB(90, 90, 90)
        For k = 0 To TAIL_POINTS - 1
            If k >= TAIL_POINTS - tailN(i) Then
                ws.cells(TBL_FIRST + k, cx).Value = tailX(i, k)
                ws.cells(TBL_FIRST + k, cx + 1).Value = tailY(i, k)
            End If
        Next k
        ws.Range(ws.cells(TBL_FIRST, cx), ws.cells(TBL_FIRST + TAIL_POINTS - 1, cx + 1)).NumberFormat = "0.00"
        ws.Range(ws.cells(TBL_FIRST, cx), ws.cells(TBL_FIRST + TAIL_POINTS - 1, cx + 1)).Font.Color = RGB(90, 90, 90)
        ws.Columns(cx).ColumnWidth = 7: ws.Columns(cx + 1).ColumnWidth = 7
    Next i

    ' ---- TRAIL: weekly RS-RATIO changes, one row per ETF, column sparkline in I ----
    Dim trCol As Long: trCol = DataCol(ws) + 2 * n + 1
    With ws.cells(TBL_HDR - 1, trCol)
        .Value = "TRAIL (weekly dRS-RATIO, sparkline source)": .Font.Color = RGB(90, 90, 90): .Font.Size = 8
    End With
    For i = 0 To n - 1
        For k = 1 To TAIL_POINTS - 1
            If k >= TAIL_POINTS - tailN(i) + 1 Then
                ws.cells(TBL_FIRST + i, trCol + k - 1).Value = tailX(i, k) - tailX(i, k - 1)
            End If
        Next k
        ws.Columns(trCol + i).ColumnWidth = 5
    Next i
    With ws.Range(ws.cells(TBL_FIRST, trCol), ws.cells(TBL_FIRST + n - 1, trCol + TAIL_POINTS - 2))
        .NumberFormat = "0.00": .Font.Color = RGB(90, 90, 90)
    End With
    Dim sg As SparklineGroup
    Set sg = ws.Range(ws.cells(TBL_FIRST, 9), ws.cells(TBL_FIRST + n - 1, 9)).SparklineGroups.Add( _
        Type:=xlSparkColumn, SourceData:=ws.Range(ws.cells(TBL_FIRST, trCol), ws.cells(TBL_FIRST + n - 1, trCol + TAIL_POINTS - 2)).Address(False, False))
    sg.SeriesColor.Color = RGB(220, 80, 80)          ' up = red (TW convention, same as the weight bars)
    sg.Points.Negative.Visible = True
    sg.Points.Negative.Color.Color = RGB(80, 200, 120)
    sg.Axes.Horizontal.Axis.Visible = False
    sg.DisplayBlanksAs = xlNotPlotted

    ' ---- chart ----
    Call DrawRrgChart(ws, tickers, tailX, tailY, tailN, n, xMin - 1.5, xMax + 1.5, yMin - 1.5, yMax + 1.5, asOf)
    Call DrawFlowChart(ws, tickers, fPchg, fCmf, fObv, fSig, fOk, n, asOf)

    Call NavAdd(ws, navCode)
    Call SetFocusMark(ws, "")
    Call SetMark(ws, SORT_MARK, "")
    Call EnsureSheetCode(ws)
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    Application.EnableCancelKey = prevCancel
    Call NavNotify("RRG rebuilt: " & n & " " & unitNm & " vs " & gBench & " as of " & Format(asOf, "yyyy/mm/dd") & _
                   IIf(isComp, "  (" & Format((Timer - t0) / 60, "0") & " min)", "") & _
                   IIf(nDropped > 0, "  .  dropped (no member >= NT$" & Format(TW_MIN_CAP / 100000000#, "0") & "e): " & dropped, ""))
    Exit Sub
Fail:
    Dim failMsg As String: failMsg = Err.Description
    Dim wasCancel As Boolean: wasCancel = (Err.Number = 18)
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    Application.EnableCancelKey = prevCancel
    On Error Resume Next
    Call NavAdd(ws, navCode)
    If wasCancel Then
        Call NavNotify("RRG stopped (Esc) - finished industries are cached, run " & navCode & "! again to continue", True)
    Else
        Call NavNotify("RRG failed: " & failMsg, True)
    End If
End Sub

' 100 * x / SMA(x, w); NaN unless the whole window is valid (pandas min_periods = window)
Private Sub RollingRatio(ByRef src() As Double, ByRef dst() As Double, ByVal w As Long)
    Dim i As Long, k As Long, s As Double, ok As Boolean
    For i = LBound(src) To UBound(src)
        dst(i) = NAN
        If i - LBound(src) + 1 >= w Then
            s = 0: ok = True
            For k = i - w + 1 To i
                If src(k) = NAN Then ok = False: Exit For
                s = s + src(k)
            Next k
            If ok And s <> 0 Then dst(i) = 100 * src(i) / (s / w)
        End If
    Next i
End Sub

' pandas Series.ewm(span=n).mean() with adjust=True (the default): weighted
' average with weights (1-a)^k, a = 2/(n+1).  Leading NaNs stay NaN.
Private Sub EwmAdjust(ByRef src() As Double, ByRef dst() As Double, ByVal spanN As Long)
    Dim a As Double: a = 2 / (spanN + 1)
    Dim num As Double, den As Double, started As Boolean
    Dim i As Long
    For i = LBound(src) To UBound(src)
        If src(i) = NAN Then
            dst(i) = NAN
            If started Then num = num * (1 - a): den = den * (1 - a)   ' ignore_na=False: the gap still decays the weights
        Else
            num = num * (1 - a) + src(i)
            den = den * (1 - a) + 1
            started = True
            dst(i) = num / den
        End If
    Next i
End Sub

Private Sub DrawRrgChart(ws As Worksheet, tickers() As String, tailX() As Double, tailY() As Double, tailN() As Long, _
                         ByVal n As Long, ByVal x0 As Double, ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double, ByVal asOf As Date)
    Dim topRow As Long: topRow = TBL_HDR + NavOffset(ws)
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(ws.Columns(CHART_COL).Left, ws.Rows(topRow).Top, gChartW, gChartH)
    co.Name = "RRG_MAIN"
    co.Placement = xlMove
    Dim ch As Chart: Set ch = co.Chart
    ch.ChartType = xlXYScatterLines
    Do While ch.SeriesCollection.count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.HasLegend = False
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(8, 8, 8)
    ch.PlotArea.Format.Line.Visible = msoFalse
    ch.HasTitle = True
    ch.ChartTitle.Text = "RRG  vs " & gBench & "   as of " & Format(asOf, "yyyy/mm/dd")
    With ch.ChartTitle.Format.TextFrame2.TextRange.Font
        .Name = PageFont(ws): .NameFarEast = PageFont(ws): .Size = 10: .Bold = msoTrue: .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    Dim i As Long, k As Long
    For i = 0 To n - 1
        If tailN(i) > 0 Then
            Dim cx As Long: cx = DataCol(ws) + 2 * i
            Dim firstRow As Long: firstRow = TBL_FIRST + NavOffset(ws) + (TAIL_POINTS - tailN(i))
            Dim lastRow As Long: lastRow = TBL_FIRST + NavOffset(ws) + TAIL_POINTS - 1
            Dim s As Series: Set s = ch.SeriesCollection.NewSeries
            s.Name = tickers(i)
            s.XValues = ws.Range(ws.cells(firstRow, cx), ws.cells(lastRow, cx))
            s.Values = ws.Range(ws.cells(firstRow, cx + 1), ws.cells(lastRow, cx + 1))
            Call StyleSeries(ws, s, QuadColor(Quadrant(tailX(i, TAIL_POINTS - 1), tailY(i, TAIL_POINTS - 1))), tailN(i), 0)
        End If
    Next i

    Dim ax As Axis
    Set ax = ch.Axes(xlCategory)
    ax.MinimumScale = x0: ax.MaximumScale = x1
    ax.CrossesAt = 100
    ax.HasMajorGridlines = False
    ax.TickLabelPosition = xlTickLabelPositionLow
    ax.Format.Line.ForeColor.RGB = RGB(110, 110, 110)
    ax.TickLabels.Font.Name = PageFont(ws): ax.TickLabels.Font.Size = 8: ax.TickLabels.Font.Color = RGB(150, 150, 150)
    ax.TickLabels.NumberFormat = "0"
    ax.HasTitle = True: ax.AxisTitle.Text = "RS-RATIO"
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = PageFont(ws)
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.NameFarEast = PageFont(ws)
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 8
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(180, 180, 180)
    Set ax = ch.Axes(xlValue)
    ax.MinimumScale = y0: ax.MaximumScale = y1
    ax.CrossesAt = 100
    ax.HasMajorGridlines = False
    ax.TickLabelPosition = xlTickLabelPositionLow
    ax.Format.Line.ForeColor.RGB = RGB(110, 110, 110)
    ax.TickLabels.Font.Name = PageFont(ws): ax.TickLabels.Font.Size = 8: ax.TickLabels.Font.Color = RGB(150, 150, 150)
    ax.TickLabels.NumberFormat = "0"
    ax.HasTitle = True: ax.AxisTitle.Text = "RS-MOMENTUM"
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = PageFont(ws)
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.NameFarEast = PageFont(ws)
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 8
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(180, 180, 180)

    ' quadrant labels in the four corners of the plot area
    Dim pl As Double, pt As Double, pw As Double, ph As Double
    pl = ch.PlotArea.InsideLeft: pt = ch.PlotArea.InsideTop
    pw = ch.PlotArea.InsideWidth: ph = ch.PlotArea.InsideHeight
    Call QuadLabel(ch, "LEADING", pl + pw - 78, pt + 4, msoAnchorTop, "Consolas")
    Call QuadLabel(ch, "WEAKENING", pl + pw - 78, pt + ph - 18, msoAnchorBottom, "Consolas")
    Call QuadLabel(ch, "LAGGING", pl + 4, pt + ph - 18, msoAnchorBottom, "Consolas")
    Call QuadLabel(ch, "IMPROVING", pl + 4, pt + 4, msoAnchorTop, "Consolas")
End Sub

' ----------------------------------------------------------------
' Series look.  mode 0 = normal (quadrant colour, ticker label on the last
' dot), 1 = dimmed (dark grey, no label), 2 = focused (thicker, every dot
' dated from the Z column).  Per-point fill transparency would reset the
' marker colour to the theme's, so fading is done by blending towards black.
Private Sub StyleSeries(ws As Worksheet, s As Series, ByVal qc As Long, ByVal np As Long, ByVal mode As Long)
    Dim k As Long
    Dim baseCol As Long: baseCol = IIf(mode = 1, DIM_GREY, qc)
    s.Format.Line.ForeColor.RGB = baseCol
    s.Format.Line.Weight = IIf(mode >= 2, 2, 1)
    s.Format.Line.Transparency = IIf(mode >= 2, 0.15, 0.5)
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 4
    s.MarkerBackgroundColor = baseCol
    s.MarkerForegroundColor = baseCol
    s.HasDataLabels = False
    For k = 1 To np - 1
        Dim fade As Double: fade = 0.25 + 0.75 * (k - 1) / (np - 1)
        s.Points(k).MarkerSize = IIf(mode >= 2, 5, 3)
        s.Points(k).MarkerBackgroundColor = IIf(mode = 1, DIM_GREY, Dim2(qc, fade))
        s.Points(k).MarkerForegroundColor = IIf(mode = 1, DIM_GREY, Dim2(qc, fade))
    Next k
    With s.Points(np)
        .MarkerSize = IIf(mode = 1, 6, IIf(mode >= 2, 11, 9))
        .MarkerBackgroundColor = baseCol
        .MarkerForegroundColor = baseCol
    End With
    Dim disp As String: disp = DisplayName(ws, s.Name)
    If mode = 0 Then
        Call PointLabel(s.Points(np), disp, RGB(210, 210, 210), IIf(ws.Name = IND_SHEET, 7, 8), PageFont(ws))
    ElseIf mode = 3 Then
        Call PointLabel(s.Points(np), disp, RGB(255, 255, 255), 9, PageFont(ws))
    ElseIf mode = 2 Then
        Dim off As Long: off = NavOffset(ws)
        For k = 1 To np
            Dim dr As Long: dr = TBL_FIRST + off + TAIL_POINTS - np + k - 1
            Dim txt As String: txt = Format(ws.cells(dr, DateCol(ws)).Value, "mm/dd")
            If k = np Then txt = disp & " " & txt
            Call PointLabel(s.Points(k), txt, IIf(k = np, RGB(255, 255, 255), RGB(190, 190, 190)), IIf(k = np, 9, 7), PageFont(ws))
        Next k
    End If
End Sub

Private Sub PointLabel(p As Point, ByVal txt As String, ByVal col As Long, ByVal sz As Long, ByVal fnt As String, Optional ByVal pos As Long = xlLabelPositionRight)
    p.HasDataLabel = True
    With p.DataLabel
        ' turning every Show* flag off deletes the label (and .Text then
        ' fails with "invalid parameter"), so keep the series name on and
        ' overwrite the text
        .ShowValue = False
        .ShowCategoryName = False
        .ShowSeriesName = True
        .Text = txt
        .Position = pos
        ' CJK glyphs take the East Asian font, not .Name - set both
        .Format.TextFrame2.TextRange.Font.Name = fnt
        .Format.TextFrame2.TextRange.Font.NameFarEast = fnt
        .Format.TextFrame2.TextRange.Font.Size = sz
        .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = col
    End With
End Sub

' Table colours for every ETF row.  focusTk = "" -> normal look; otherwise
' that row gets the orange background and every other row goes dim.
Private Sub PaintTableRows(ws As Worksheet, ByVal focusTk As String)
    Dim lc As Long: lc = NavLeft(ws)             ' blank column A (nav bar) shifts every table column
    Dim off As Long: off = NavOffset(ws)
    Dim r As Long, i As Long
    i = 0
    r = TBL_FIRST + off
    Do While ws.cells(r, 1 + lc).Value <> ""
        Dim rng As Range: Set rng = ws.Range(ws.cells(r, 1 + lc), ws.cells(r, TBL_NCOL + lc))
        Dim tk As String: tk = CStr(ws.cells(r, 1 + lc).Value)
        Dim q As String: q = CStr(ws.cells(r, 6 + lc).Value)
        rng.Font.Bold = False
        If focusTk <> "" And InSet(focusTk, tk) Then
            rng.Interior.Color = RR4_ACCENT
            rng.Font.Color = RGB(0, 0, 0)
            rng.Font.Bold = True
        ElseIf focusTk <> "" Then
            rng.Interior.Color = IIf(i Mod 2 = 0, RGB(8, 8, 8), RGB(14, 14, 14))
            rng.Font.Color = RGB(75, 75, 75)
        Else
            rng.Interior.Color = IIf(i Mod 2 = 0, RGB(8, 8, 8), RGB(14, 14, 14))
            rng.Font.Color = RGB(221, 221, 221)
            ws.cells(r, 1 + lc).Font.Color = RR4_ACCENT: ws.cells(r, 1 + lc).Font.Bold = True
            ws.cells(r, 3 + lc).Font.Color = RGB(150, 150, 150)
            If q = "NO DATA" Then
                ws.cells(r, 6 + lc).Font.Color = RGB(120, 120, 120)
            Else
                ws.cells(r, 6 + lc).Font.Color = QuadColor(q): ws.cells(r, 6 + lc).Font.Bold = True
            End If
            If ws.cells(r, 7 + lc).Value <> "" Then
                ws.cells(r, 7 + lc).Font.Color = IIf(ws.cells(r, 7 + lc).Value >= 0, RGB(220, 80, 80), RGB(80, 200, 120))
                ws.cells(r, 8 + lc).Font.Color = IIf(ws.cells(r, 8 + lc).Value >= 0, RGB(220, 80, 80), RGB(80, 200, 120))
            End If
            ws.cells(r, 10 + lc).Font.Color = RGB(120, 120, 120)
            If ws.cells(r, 11 + lc).Value <> "" Then
                ws.cells(r, 11 + lc).Font.Color = IIf(ws.cells(r, 11 + lc).Value >= 0, RGB(220, 80, 80), RGB(80, 200, 120))
                ws.cells(r, 12 + lc).Font.Color = IIf(ws.cells(r, 12 + lc).Value >= 0, RGB(220, 80, 80), RGB(80, 200, 120))
                ws.cells(r, 13 + lc).Font.Color = IIf(ws.cells(r, 13 + lc).Value >= 0, RGB(220, 80, 80), RGB(80, 200, 120))
            End If
            ws.cells(r, 14 + lc).Font.Color = SigColor(CStr(ws.cells(r, 14 + lc).Value)): ws.cells(r, 14 + lc).Font.Bold = True
        End If
        i = i + 1: r = r + 1
    Loop
End Sub

' ----------------------------------------------------------------
' Focus: "" shows every ETF, otherwise only the tickers in tk (comma-
' separated set) stay lit - one ticker also gets its dates, several only
' their names.
Public Sub RrgFocus(ws As Worksheet, ByVal tk As String)
    Dim lc As Long: lc = NavLeft(ws)             ' blank column A (nav bar) shifts every table column
    tk = NormSet(tk)
    Dim litMode As Long: litMode = IIf(SetCount(tk) = 1, 2, 3)
    Dim co As ChartObject
    On Error Resume Next
    Set co = ws.ChartObjects("RRG_MAIN")
    On Error GoTo 0
    If co Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    ' quadrant per ticker from the table (the chart series carry no colour memory)
    Dim off As Long: off = NavOffset(ws)
    Dim quadOf As Object: Set quadOf = CreateObject("Scripting.Dictionary")
    Dim r As Long: r = TBL_FIRST + off
    Do While ws.cells(r, 1 + lc).Value <> ""
        quadOf(CStr(ws.cells(r, 1 + lc).Value)) = CStr(ws.cells(r, 6 + lc).Value)
        r = r + 1
    Loop
    Dim s As Series
    For Each s In co.Chart.SeriesCollection
        Dim mode As Long
        If tk = "" Then mode = 0 Else mode = IIf(InSet(tk, s.Name), litMode, 1)
        Call StyleSeries(ws, s, QuadColor(CStr(quadOf(s.Name))), s.Points.count, mode)
    Next s
    Dim cf As ChartObject
    On Error Resume Next
    Set cf = ws.ChartObjects("RRG_FLOW")
    On Error GoTo 0
    If Not cf Is Nothing Then
        Dim sigOf As Object: Set sigOf = CreateObject("Scripting.Dictionary")
        Dim obvOf As Object: Set obvOf = CreateObject("Scripting.Dictionary")
        r = TBL_FIRST + off
        Do While ws.cells(r, 1 + lc).Value <> ""
            sigOf(CStr(ws.cells(r, 1 + lc).Value)) = CStr(ws.cells(r, 14 + lc).Value)
            obvOf(CStr(ws.cells(r, 1 + lc).Value)) = Val(ws.cells(r, 13 + lc).Value)
            r = r + 1
        Loop
        For Each s In cf.Chart.SeriesCollection
            If Left(s.Name, 1) <> "_" Then
                If tk = "" Then mode = 0 Else mode = IIf(InSet(tk, s.Name), litMode, 1)
                Call StyleFlowSeries(ws, s, SigColor(CStr(sigOf(s.Name))), CDbl(obvOf(s.Name)), mode)
            End If
        Next s
    End If
    Call PaintTableRows(ws, tk)
    Call SetFocusMark(ws, tk)
    Application.ScreenUpdating = True
    If tk = "" Then
        Call NavNotify("RRG: showing all")
    Else
        Call NavNotify("RRG: focus " & Replace(tk, ",", " + ") & IIf(SetCount(tk) = 1, " (" & quadOf(tk) & ")", "") & _
                       "  -  double-click a ticker to add / remove it, the title to show all")
    End If
End Sub

' Same, by sheet name - callable from the Immediate window / Application.Run.
Public Sub RrgFocusByName(ByVal tk As String, Optional ByVal sheetNm As String = ETF_SHEET)
    Call RrgFocus(ThisWorkbook.Sheets(sheetNm), UCase(Trim(tk)))
End Sub

' Sheet double-click handler (called from SheetRRG_Code.txt).
Public Sub RrgDoubleClick(ws As Worksheet, ByVal Target As Range, ByRef Cancel As Boolean)
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim tcol As Long: tcol = Target.Column - lc                 ' table column (1 = ticker)
    If Target.Row = 1 + off And tcol = 1 Then                   ' page title: reset everything
        Cancel = True
        Call RrgFocus(ws, "")
        Call RrgSort(ws, 0)
    ElseIf Target.Row = TBL_HDR + off And tcol >= 1 And tcol <= TBL_NCOL Then
        Cancel = True
        Call RrgSort(ws, tcol)
    ElseIf tcol = 1 And Target.Row > TBL_HDR + off Then
        Dim tk As String: tk = Trim(CStr(Target.Value))
        If tk = "" Or ws.cells(Target.Row, 6 + lc).Value = "" Then Exit Sub    ' below the table
        Cancel = True
        Call RrgFocus(ws, ToggleSet(GetFocusMark(ws), tk))
    End If
End Sub

Private Sub SetFocusMark(ws As Worksheet, ByVal tk As String)
    Call SetMark(ws, FOCUS_MARK, tk)
End Sub

Private Function GetFocusMark(ws As Worksheet) As String
    GetFocusMark = GetMark(ws, FOCUS_MARK)
End Function

Private Sub SetMark(ws As Worksheet, ByVal nm As String, ByVal v As String)
    On Error Resume Next
    ws.Names(nm).Delete
    On Error GoTo 0
    ws.Names.Add Name:=nm, RefersTo:="=""" & v & """", Visible:=False
End Sub

Private Function GetMark(ws As Worksheet, ByVal nm As String) As String
    On Error Resume Next
    Dim v As String: v = ws.Names(nm).RefersTo      ' ="SMH"
    On Error GoTo 0
    GetMark = Replace(Replace(v, "=", ""), """", "")
End Function

' Write the sheet's event code (a copy of RR4/SheetRRG_Code.txt) into the
' document module if it is not there yet.  The RRG sheet is created by
' BuildRRG, so the code cannot be injected ahead of time.
Private Sub EnsureSheetCode(ws As Worksheet)
    On Error GoTo Skip
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 Then
            If comp.Properties("Name").Value = ws.Name Then
                Dim cm As Object: Set cm = comp.CodeModule
                If cm.CountOfLines > 0 Then
                    If InStr(cm.Lines(1, cm.CountOfLines), "RrgDoubleClick") > 0 Then Exit Sub
                    cm.DeleteLines 1, cm.CountOfLines
                End If
                cm.AddFromString "Option Explicit" & vbCrLf & vbCrLf & _
                    "' RRG page: double-click a ticker in column A to focus it (see RRG.bas)" & vbCrLf & _
                    "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & vbCrLf & _
                    "    Call RrgDoubleClick(Me, Target, Cancel)" & vbCrLf & _
                    "End Sub" & vbCrLf
                Exit Sub
            End If
        End If
    Next comp
    Exit Sub
Skip:
    Call NavNotify("RRG built, but the double-click code could not be written (enable Trust access to the VBA project object model, or paste RR4/SheetRRG_Code.txt)", True)
End Sub

' ----------------------------------------------------------------
' money_flow_ranking.compute_ticker_signal() on one ticker's rows.
Private Function MoneyFlow(ByRef c() As Double, ByRef h() As Double, ByRef l() As Double, ByRef v() As Double, ByVal nt As Long, _
                           ByRef pchg As Double, ByRef cmf As Double, ByRef obvT As Double, ByRef sig As String, ByRef note As String) As Boolean
    If nt < CMF_WINDOW + FLOW_LOOKBACK Then Exit Function
    Dim k As Long
    ' CMF over the last CMF_WINDOW rows (pandas rolling: a zero range -> NaN -> flat)
    Dim mfv As Double, vol As Double, bad As Boolean
    For k = nt - CMF_WINDOW To nt - 1
        If h(k) - l(k) = 0 Then bad = True: Exit For
        mfv = mfv + ((c(k) - l(k)) - (h(k) - c(k))) / (h(k) - l(k)) * v(k)
        vol = vol + v(k)
    Next k
    If bad Or vol = 0 Then cmf = 0 Else cmf = mfv / vol
    ' OBV change over FLOW_LOOKBACK rows, in days of average volume
    Dim dObv As Double, avgV As Double
    For k = nt - FLOW_LOOKBACK To nt - 1
        If c(k) > c(k - 1) Then dObv = dObv + v(k) Else If c(k) < c(k - 1) Then dObv = dObv - v(k)
        avgV = avgV + v(k)
    Next k
    avgV = avgV / FLOW_LOOKBACK
    If avgV = 0 Then obvT = 0 Else obvT = dObv / avgV
    pchg = c(nt - 1) / c(nt - 1 - FLOW_LOOKBACK) - 1
    sig = FlowSignal(pchg, cmf, obvT, note)
    MoneyFlow = True
End Function

' ----------------------------------------------------------------
' Money flow quadrant chart under the RRG: x = 20-day price change (table
' column J), y = CMF (column K), one single-point series per ETF so each
' can carry its own marker shape (OBV direction) and colour (signal).
Private Sub DrawFlowChart(ws As Worksheet, tickers() As String, fPchg() As Double, fCmf() As Double, fObv() As Double, _
                          fSig() As String, fOk() As Boolean, ByVal n As Long, ByVal asOf As Date)
    Dim off As Long: off = NavOffset(ws)
    Dim topPt As Double: topPt = ws.Rows(TBL_HDR + off).Top + gChartH + 12
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(ws.Columns(CHART_COL).Left, topPt, gChartW, gChartH)
    co.Name = "RRG_FLOW"
    co.Placement = xlMove
    Dim ch As Chart: Set ch = co.Chart
    ch.ChartType = xlXYScatter
    Do While ch.SeriesCollection.count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.HasLegend = False
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(8, 8, 8)
    ch.PlotArea.Format.Line.Visible = msoFalse
    ch.HasTitle = True
    ch.ChartTitle.Text = "MONEY FLOW  price change vs Chaikin Money Flow   as of " & Format(asOf, "yyyy/mm/dd") & _
                         "   (CMF " & CMF_WINDOW & "d, lookback " & FLOW_LOOKBACK & "d)"
    With ch.ChartTitle.Format.TextFrame2.TextRange.Font
        .Name = PageFont(ws): .NameFarEast = PageFont(ws): .Size = 10: .Bold = msoTrue: .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    ' axis range (money_flow_chart.build_figure: 15% pad, at least 2% / 0.02)
    Dim i As Long, maxX As Double, maxY As Double, x0 As Double, x1 As Double, y0 As Double, y1 As Double
    Dim anyOk As Boolean
    For i = 0 To n - 1
        If fOk(i) Then
            anyOk = True
            If Abs(fPchg(i)) > maxX Then maxX = Abs(fPchg(i))
            If Abs(fCmf(i)) > maxY Then maxY = Abs(fCmf(i))
            If fPchg(i) < x0 Then x0 = fPchg(i)
            If fPchg(i) > x1 Then x1 = fPchg(i)
            If fCmf(i) < y0 Then y0 = fCmf(i)
            If fCmf(i) > y1 Then y1 = fCmf(i)
        End If
    Next i
    If Not anyOk Then Exit Sub
    Dim padX As Double: padX = maxX * 0.15: If padX < 0.02 Then padX = 0.02
    Dim padY As Double: padY = maxY * 0.15: If padY < 0.02 Then padY = 0.02
    x0 = x0 - padX: x1 = x1 + padX: y0 = y0 - padY: y1 = y1 + padY
    If y1 < CMF_DEADZONE + padY Then y1 = CMF_DEADZONE + padY
    If y0 > -CMF_DEADZONE - padY Then y0 = -CMF_DEADZONE - padY

    ' dead-zone lines at +/- CMF_DEADZONE: two helper series from the data block
    Dim dzRow As Long: dzRow = TBL_FIRST + off + TAIL_POINTS + 2
    Dim dzCol As Long: dzCol = DataCol(ws)
    ws.cells(dzRow - 1, dzCol).Value = "dead zone": ws.cells(dzRow - 1, dzCol).Font.Color = RGB(90, 90, 90)
    ws.cells(dzRow, dzCol).Value = x0: ws.cells(dzRow, dzCol + 1).Value = CMF_DEADZONE
    ws.cells(dzRow + 1, dzCol).Value = x1: ws.cells(dzRow + 1, dzCol + 1).Value = CMF_DEADZONE
    ws.cells(dzRow + 2, dzCol).Value = x0: ws.cells(dzRow + 2, dzCol + 1).Value = -CMF_DEADZONE
    ws.cells(dzRow + 3, dzCol).Value = x1: ws.cells(dzRow + 3, dzCol + 1).Value = -CMF_DEADZONE
    ws.Range(ws.cells(dzRow, dzCol), ws.cells(dzRow + 3, dzCol + 1)).Font.Color = RGB(90, 90, 90)
    ws.Range(ws.cells(dzRow, dzCol), ws.cells(dzRow + 3, dzCol + 1)).NumberFormat = "0.000"
    Dim d As Long
    For d = 0 To 1
        Dim dz As Series: Set dz = ch.SeriesCollection.NewSeries
        dz.Name = "_dz" & d
        dz.XValues = ws.Range(ws.cells(dzRow + 2 * d, dzCol), ws.cells(dzRow + 2 * d + 1, dzCol))
        dz.Values = ws.Range(ws.cells(dzRow + 2 * d, dzCol + 1), ws.cells(dzRow + 2 * d + 1, dzCol + 1))
        dz.ChartType = xlXYScatterLinesNoMarkers
        dz.Format.Line.ForeColor.RGB = RGB(80, 80, 80)
        dz.Format.Line.Weight = 0.75
        dz.Format.Line.DashStyle = msoLineDash
    Next d

    ' one point per ETF
    For i = 0 To n - 1
        If fOk(i) Then
            Dim r As Long: r = TBL_FIRST + off + i
            Dim s As Series: Set s = ch.SeriesCollection.NewSeries
            s.Name = tickers(i)
            s.XValues = ws.Range(ws.cells(r, 11), ws.cells(r, 11))
            s.Values = ws.Range(ws.cells(r, 12), ws.cells(r, 12))
            s.ChartType = xlXYScatter
            Call StyleFlowSeries(ws, s, SigColor(fSig(i)), fObv(i), 0)
        End If
    Next i

    Dim ax As Axis
    Set ax = ch.Axes(xlCategory)
    ax.MinimumScale = x0: ax.MaximumScale = x1
    ax.CrossesAt = 0
    ax.HasMajorGridlines = False
    ax.TickLabelPosition = xlTickLabelPositionLow
    ax.Format.Line.ForeColor.RGB = RGB(110, 110, 110)
    ax.TickLabels.Font.Name = PageFont(ws): ax.TickLabels.Font.Size = 8: ax.TickLabels.Font.Color = RGB(150, 150, 150)
    ax.TickLabels.NumberFormat = "0%"
    ax.HasTitle = True: ax.AxisTitle.Text = "PRICE CHANGE " & FLOW_LOOKBACK & "D"
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = PageFont(ws)
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.NameFarEast = PageFont(ws)
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 8
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(180, 180, 180)
    Set ax = ch.Axes(xlValue)
    ax.MinimumScale = y0: ax.MaximumScale = y1
    ax.CrossesAt = 0
    ax.HasMajorGridlines = False
    ax.TickLabelPosition = xlTickLabelPositionLow
    ax.Format.Line.ForeColor.RGB = RGB(110, 110, 110)
    ax.TickLabels.Font.Name = PageFont(ws): ax.TickLabels.Font.Size = 8: ax.TickLabels.Font.Color = RGB(150, 150, 150)
    ax.TickLabels.NumberFormat = "0.00"
    ax.HasTitle = True: ax.AxisTitle.Text = "CHAIKIN MONEY FLOW"
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = PageFont(ws)
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.NameFarEast = PageFont(ws)
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 8
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(180, 180, 180)

    Dim pl As Double, pt As Double, pw As Double, ph As Double
    pl = ch.PlotArea.InsideLeft: pt = ch.PlotArea.InsideTop
    pw = ch.PlotArea.InsideWidth: ph = ch.PlotArea.InsideHeight
    Call CornerLabel(ch, "CONFIRMED INFLOW", SigColor("INFLOW"), pl + pw - 124, pt + 4, True, "Consolas")
    Call CornerLabel(ch, "BEARISH DIVERGENCE", SigColor("BEAR DIV"), pl + pw - 124, pt + ph - 18, True, "Consolas")
    Call CornerLabel(ch, "CONFIRMED OUTFLOW", SigColor("OUTFLOW"), pl + 4, pt + ph - 18, False, "Consolas")
    Call CornerLabel(ch, "BULLISH DIVERGENCE", SigColor("BULL DIV"), pl + 4, pt + 4, False, "Consolas")
End Sub

' mode 0 normal / 1 dimmed / 2 focused, same meaning as StyleSeries.
' Marker shape = OBV direction (triangle up, diamond down, circle flat).
Private Sub StyleFlowSeries(ws As Worksheet, s As Series, ByVal col As Long, ByVal obvT As Double, ByVal mode As Long)
    Dim c As Long: c = IIf(mode = 1, DIM_GREY, col)
    Dim disp As String: disp = DisplayName(ws, s.Name)
    If obvT > OBV_DEADZONE Then
        s.MarkerStyle = xlMarkerStyleTriangle
    ElseIf obvT < -OBV_DEADZONE Then
        s.MarkerStyle = xlMarkerStyleDiamond
    Else
        s.MarkerStyle = xlMarkerStyleCircle
    End If
    s.MarkerSize = IIf(mode = 1, 6, IIf(mode >= 2, 14, 9))
    s.MarkerBackgroundColor = c
    s.MarkerForegroundColor = c
    s.HasDataLabels = False
    If mode = 0 Then
        Call PointLabel(s.Points(1), disp, RGB(210, 210, 210), IIf(ws.Name = IND_SHEET, 7, 8), PageFont(ws), xlLabelPositionAbove)
    ElseIf mode >= 2 Then
        Call PointLabel(s.Points(1), disp, RGB(255, 255, 255), 10, PageFont(ws), xlLabelPositionAbove)
    End If
End Sub

' What a series is called on the charts: the ticker on the ETF page, the
' industry name (table column B) on the industry page - column A there is
' the plate code, which is the series / focus key but says nothing.
Private Function DisplayName(ws As Worksheet, ByVal key As String) As String
    Dim lc As Long: lc = NavLeft(ws)             ' blank column A (nav bar) shifts every table column
    DisplayName = key
    If ws.Name <> IND_SHEET Then Exit Function
    Dim r As Long: r = TBL_FIRST + NavOffset(ws)
    Do While ws.cells(r, 1 + lc).Value <> ""
        If StrComp(CStr(ws.cells(r, 1 + lc).Value), key, vbTextCompare) = 0 Then
            If ws.cells(r, 2 + lc).Value <> "" Then DisplayName = CStr(ws.cells(r, 2 + lc).Value)
            Exit Function
        End If
        r = r + 1
    Loop
End Function

Private Sub CornerLabel(ch As Chart, ByVal txt As String, ByVal col As Long, ByVal l As Double, ByVal t As Double, ByVal alignRight As Boolean, ByVal fnt As String)
    Dim shp As Shape
    Set shp = ch.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, 120, 14)
    shp.Name = "RRG_Q_" & txt
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Name = fnt
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = col
        .TextRange.ParagraphFormat.Alignment = IIf(alignRight, msoAlignRight, msoAlignLeft)
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub

' ----------------------------------------------------------------
' SORT: double-click a column header.  Same header again flips the
' direction; the page title restores the build order (SEQ column Y).
' The table rows (A:N), the TRAIL block and the SEQ column move together;
' the tail data block is per-ticker columns so it is untouched, and the
' RRG_FLOW series (which point at table cells) are re-bound by ticker.
Public Sub RrgSort(ws As Worksheet, ByVal col As Long)
    Dim lc As Long: lc = NavLeft(ws)             ' blank column A (nav bar) shifts every table column
    Dim off As Long: off = NavOffset(ws)
    Dim r0 As Long: r0 = TBL_FIRST + off
    Dim n As Long: n = 0
    Do While ws.cells(r0 + n, 1 + lc).Value <> ""
        n = n + 1
    Loop
    If n < 2 Then Exit Sub
    Application.ScreenUpdating = False
    Dim trCol As Long: trCol = DataCol(ws) + 2 * n + 1
    Dim tw As Long: tw = TAIL_POINTS - 1

    ' --- direction: same column again flips it ---
    Dim prev As String: prev = GetMark(ws, SORT_MARK)         ' "col|dir"
    Dim dir As Long
    Dim isText As Boolean: isText = (col = 1 Or col = 2 Or col = 3 Or col = 6 Or col = TBL_NCOL)
    If col = 0 Then
        dir = 1
    ElseIf Split(prev & "|", "|")(0) = CStr(col) Then
        dir = -CLng(Val(Split(prev & "|", "|")(1)))
    Else
        dir = IIf(isText, 1, -1)                              ' numbers: biggest first
    End If

    ' --- read everything that moves ---
    Dim tbl As Variant: tbl = ws.Range(ws.cells(r0, 1 + lc), ws.cells(r0 + n - 1, TBL_NCOL + lc)).Value
    Dim trl As Variant: trl = ws.Range(ws.cells(r0, trCol), ws.cells(r0 + n - 1, trCol + tw - 1)).Value
    Dim seq As Variant: seq = ws.Range(ws.cells(r0, SeqCol(ws)), ws.cells(r0 + n - 1, SeqCol(ws))).Value
    Dim notes() As String: ReDim notes(1 To n)
    Dim i As Long, j As Long
    For i = 1 To n
        If Not ws.cells(r0 + i - 1, TBL_NCOL + lc).Comment Is Nothing Then notes(i) = ws.cells(r0 + i - 1, TBL_NCOL + lc).Comment.Text
    Next i

    ' --- order (insertion sort on n ~ 26 rows; blanks always last) ---
    Dim idx() As Long: ReDim idx(1 To n)
    For i = 1 To n: idx(i) = i: Next i
    For i = 2 To n
        Dim cur As Long: cur = idx(i)
        j = i - 1
        Do While j >= 1
            If SortBefore(tbl, seq, idx(j), cur, col, dir, isText) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = cur
    Next i

    ' --- write back in the new order ---
    Dim tbl2 As Variant: tbl2 = tbl
    Dim trl2 As Variant: trl2 = trl
    Dim seq2 As Variant: seq2 = seq
    Dim c As Long
    For i = 1 To n
        For c = 1 To TBL_NCOL: tbl2(i, c) = tbl(idx(i), c): Next c
        For c = 1 To tw: trl2(i, c) = trl(idx(i), c): Next c
        seq2(i, 1) = seq(idx(i), 1)
    Next i
    ' column I holds the sparklines - leave it alone, write A:H and J:N
    Dim lv As Variant: ReDim lv(1 To n, 1 To 8)
    Dim rv As Variant: ReDim rv(1 To n, 1 To TBL_NCOL - 9)
    For i = 1 To n
        For c = 1 To 8: lv(i, c) = tbl2(i, c): Next c
        For c = 10 To TBL_NCOL: rv(i, c - 9) = tbl2(i, c): Next c
    Next i
    ws.Range(ws.cells(r0, 1 + lc), ws.cells(r0 + n - 1, 8 + lc)).Value = lv
    ws.Range(ws.cells(r0, 10 + lc), ws.cells(r0 + n - 1, TBL_NCOL + lc)).Value = rv
    ws.Range(ws.cells(r0, trCol), ws.cells(r0 + n - 1, trCol + tw - 1)).Value = trl2
    ws.Range(ws.cells(r0, SeqCol(ws)), ws.cells(r0 + n - 1, SeqCol(ws))).Value = seq2
    ' number formats were set per cell at build time (rows without data had none)
    ws.Range(ws.cells(r0, 4 + lc), ws.cells(r0 + n - 1, 5 + lc)).NumberFormat = "0.00"
    ws.Range(ws.cells(r0, 7 + lc), ws.cells(r0 + n - 1, 8 + lc)).NumberFormat = "+0.00;-0.00;0.00"
    ws.Range(ws.cells(r0, 11 + lc), ws.cells(r0 + n - 1, 11 + lc)).NumberFormat = "+0.0%;-0.0%;0.0%"
    ws.Range(ws.cells(r0, 12 + lc), ws.cells(r0 + n - 1, 12 + lc)).NumberFormat = "+0.000;-0.000;0.000"
    ws.Range(ws.cells(r0, 13 + lc), ws.cells(r0 + n - 1, 13 + lc)).NumberFormat = "+0.0;-0.0;0.0"
    ws.Range(ws.cells(r0, TBL_NCOL + lc), ws.cells(r0 + n - 1, TBL_NCOL + lc)).ClearComments
    For i = 1 To n
        If notes(idx(i)) <> "" Then ws.cells(r0 + i - 1, TBL_NCOL + lc).AddComment notes(idx(i))
    Next i

    ' --- RRG_FLOW series point at table cells: re-bind by ticker ---
    Dim rowOf As Object: Set rowOf = CreateObject("Scripting.Dictionary")
    For i = 1 To n: rowOf(CStr(tbl2(i, 1))) = r0 + i - 1: Next i
    Dim cf As ChartObject
    On Error Resume Next
    Set cf = ws.ChartObjects("RRG_FLOW")
    On Error GoTo 0
    If Not cf Is Nothing Then
        Dim s As Series
        For Each s In cf.Chart.SeriesCollection
            If Left(s.Name, 1) <> "_" Then
                If rowOf.Exists(s.Name) Then
                    s.XValues = ws.Range(ws.cells(rowOf(s.Name), 11 + lc), ws.cells(rowOf(s.Name), 11 + lc))
                    s.Values = ws.Range(ws.cells(rowOf(s.Name), 12 + lc), ws.cells(rowOf(s.Name), 12 + lc))
                End If
            End If
        Next s
    End If

    ' --- header marker + colours ---
    Dim hr As Long: hr = TBL_HDR + off
    ws.Range(ws.cells(hr, 1 + lc), ws.cells(hr, TBL_NCOL + lc)).Font.Underline = xlUnderlineStyleNone
    If col > 0 Then ws.cells(hr, col + lc).Font.Underline = xlUnderlineStyleSingle
    Call SetMark(ws, SORT_MARK, IIf(col = 0, "", col & "|" & dir))
    Call PaintTableRows(ws, GetMark(ws, FOCUS_MARK))
    Application.ScreenUpdating = True
    If col = 0 Then
        Call NavNotify("RRG: build order restored")
    Else
        Call NavNotify("RRG: sorted by " & ws.cells(hr, col + lc).Value & IIf(dir > 0, " ascending", " descending") & "  -  double-click the header again to flip, the title to restore")
    End If
End Sub

' True when row a should come before row b.
Private Function SortBefore(ByRef tbl As Variant, ByRef seq As Variant, ByVal a As Long, ByVal b As Long, _
                            ByVal col As Long, ByVal dir As Long, ByVal isText As Boolean) As Boolean
    If col = 0 Then
        SortBefore = (Val(seq(a, 1)) <= Val(seq(b, 1)))
        Exit Function
    End If
    Dim va As Variant, vb As Variant
    va = tbl(a, col): vb = tbl(b, col)
    Dim ea As Boolean, eb As Boolean
    ea = IsEmpty(va) Or (CStr(va) = ""): eb = IsEmpty(vb) Or (CStr(vb) = "")
    If ea And eb Then SortBefore = (Val(seq(a, 1)) <= Val(seq(b, 1))): Exit Function
    If ea Then SortBefore = False: Exit Function
    If eb Then SortBefore = True: Exit Function
    Dim cmp As Long
    If isText Then
        cmp = StrComp(CStr(va), CStr(vb), vbTextCompare)
    Else
        If CDbl(va) < CDbl(vb) Then cmp = -1 Else If CDbl(va) > CDbl(vb) Then cmp = 1 Else cmp = 0
    End If
    If cmp = 0 Then SortBefore = (Val(seq(a, 1)) <= Val(seq(b, 1))) Else SortBefore = (cmp * dir < 0)
End Function

Public Sub RrgSortByName(ByVal col As Long, Optional ByVal sheetNm As String = ETF_SHEET)
    Call RrgSort(ThisWorkbook.Sheets(sheetNm), col)
End Sub

' ---- comma-separated ticker sets (the focus group) ----
Private Function NormSet(ByVal v As String) As String
    Dim p As Variant, out As String
    For Each p In Split(v, ",")
        If Trim(p) <> "" Then out = out & IIf(out = "", "", ",") & UCase(Trim(p))
    Next p
    NormSet = out
End Function

Private Function InSet(ByVal v As String, ByVal tk As String) As Boolean
    InSet = InStr("," & v & ",", "," & UCase(tk) & ",") > 0
End Function

Private Function SetCount(ByVal v As String) As Long
    If v = "" Then SetCount = 0 Else SetCount = UBound(Split(v, ",")) + 1
End Function

Private Function ToggleSet(ByVal v As String, ByVal tk As String) As String
    If InSet(v, tk) Then
        ToggleSet = NormSet(Replace("," & v & ",", "," & UCase(tk) & ",", ","))
    Else
        ToggleSet = NormSet(v & "," & tk)
    End If
End Function

' colour scaled towards black: f = 1 keeps it, f = 0.25 is a quarter as bright
Private Function Dim2(ByVal c As Long, ByVal f As Double) As Long
    Dim2 = RGB(Int((c Mod 256) * f), Int(((c \ 256) Mod 256) * f), Int(((c \ 65536) Mod 256) * f))
End Function

Private Sub QuadLabel(ch As Chart, ByVal txt As String, ByVal l As Double, ByVal t As Double, ByVal anchor As Long, ByVal fnt As String)
    Call CornerLabel(ch, txt, QuadColor(txt), l, t, (txt = "LEADING" Or txt = "WEAKENING"), fnt)
End Sub

' ================================================================
'  INDUSTRY UNIVERSE + COMPOSITES  (RI page)
' ================================================================

' Reads the IndustryMap sheet (A plateCode, B plateName, C symbol, D rank,
' E marketCap, F plateType; header row 1, one row per constituent) into
' one entry per industry, in first-seen order.  indSyms(i) / indCaps(i) are
' the industry's constituent symbols and market caps (rank order).
' Returns the number of industries (0 = sheet missing / empty).
Private Function IndustryUniverse(ByRef tickers() As String, ByRef labels() As String, ByRef groups() As String, _
                                  ByRef indSyms As Variant, ByRef indCaps As Variant, ByRef nStocks As Long) As Long
    Dim wm As Worksheet
    On Error Resume Next
    Set wm = ThisWorkbook.Sheets(IND_MAP)
    On Error GoTo 0
    If wm Is Nothing Then Exit Function
    Dim lastR As Long: lastR = wm.cells(wm.Rows.count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function
    Dim v As Variant: v = wm.Range(wm.cells(2, 1), wm.cells(lastR, 6)).Value
    Dim nr As Long: nr = UBound(v, 1)
    Dim idxOf As Object: Set idxOf = CreateObject("Scripting.Dictionary")
    Dim syms() As Variant, caps() As Variant, cnt() As Long
    ReDim tickers(0 To nr - 1): ReDim labels(0 To nr - 1): ReDim groups(0 To nr - 1)
    ReDim syms(0 To nr - 1): ReDim caps(0 To nr - 1): ReDim cnt(0 To nr - 1)
    Dim n As Long, r As Long, i As Long
    Dim tmpS() As String, tmpC() As Double
    For r = 1 To nr
        Dim code As String: code = Trim(CStr(v(r, 1)))
        Dim sym As String: sym = Trim(CStr(v(r, 3)))
        If code <> "" And sym <> "" Then
            If Not idxOf.Exists(code) Then
                idxOf(code) = n
                tickers(n) = code
                labels(n) = Trim(CStr(v(r, 2)))
                groups(n) = ""
                ReDim tmpS(0 To 0): ReDim tmpC(0 To 0)
                syms(n) = tmpS: caps(n) = tmpC: cnt(n) = 0
                n = n + 1
            End If
            i = idxOf(code)
            Dim sArr() As String: sArr = syms(i)
            Dim cArr() As Double: cArr = caps(i)
            If cnt(i) > 0 Then ReDim Preserve sArr(0 To cnt(i)): ReDim Preserve cArr(0 To cnt(i))
            sArr(cnt(i)) = sym
            cArr(cnt(i)) = IIf(IsNumeric(v(r, 5)), CDbl(v(r, 5)), 0)
            syms(i) = sArr: caps(i) = cArr
            cnt(i) = cnt(i) + 1
            nStocks = nStocks + 1
        End If
    Next r
    If n = 0 Then Exit Function
    ReDim Preserve tickers(0 To n - 1): ReDim Preserve labels(0 To n - 1): ReDim Preserve groups(0 To n - 1)
    ReDim Preserve syms(0 To n - 1): ReDim Preserve caps(0 To n - 1)
    indSyms = syms: indCaps = caps
    IndustryUniverse = n
End Function

' moomoo symbol -> Yahoo symbol (BRK.B -> BRK-B).
Private Function YahooSymbol(ByVal sym As String) As String
    sym = Trim(sym)
    If Right(UCase(sym), 3) = ".TW" Or Right(UCase(sym), 4) = ".TWO" Then
        YahooSymbol = UCase(sym)                     ' exchange suffix stays
    Else
        YahooSymbol = Replace(sym, ".", "-")         ' BRK.B -> BRK-B
    End If
End Function

' One industry's composite OHLCV on the benchmark's days.  Served from the
' IndustryPx cache when a row for this code with the same as-of day exists;
' otherwise every constituent is fetched and the chain-linked index built
' (see the module header), then cached.  Returns the number of points.
' okC / failC accumulate the constituents with / without data over the run.
' capIsShares (TW groups): caps are share counts, w = shares x last close,
' members below TW_MIN_CAP are skipped and nm is rewritten to the largest
' counted members.
Private Function IndustryComposite(ByVal code As String, ByRef nm As String, ByVal syms As Variant, ByVal caps As Variant, _
                                   ByVal capIsShares As Boolean, _
                                   ByRef bDays() As Long, ByVal nb As Long, ByVal bIdx As Object, _
                                   ByVal iNo As Long, ByVal nInd As Long, ByVal t0 As Double, _
                                   ByRef days() As Long, ByRef c() As Double, ByRef h() As Double, ByRef l() As Double, ByRef v() As Double, _
                                   ByRef okC As Long, ByRef failC As Long, ByRef okThis As Long) As Long
    Dim asOf As Long: asOf = bDays(nb - 1)
    Dim nOk As Long, nFail As Long, nExcl As Long, npts As Long
    Dim ckey As String: ckey = code
    If capIsShares Then ckey = code & "|cap" & Format(TW_MIN_CAP / 100000000#, "0") & "e"
    Dim cLbl As String
    npts = CacheRead(ckey, asOf, days, c, h, l, v, nOk, nFail, nExcl, cLbl)
    If npts <> 0 Then                                   ' -1 = cached, but nothing counted
        okC = okC + nOk: failC = failC + nFail: okThis = nOk
        gExclC = gExclC + nExcl
        If capIsShares And cLbl <> "" Then nm = cLbl
        IndustryComposite = IIf(npts > 0, npts, 0)
        Exit Function
    End If

    Dim m As Long: m = UBound(syms) + 1
    Dim incCode() As String, incCap() As Double, nInc As Long   ' counted members, for the column B label
    ReDim incCode(0 To m): ReDim incCap(0 To m)
    Dim sumW() As Double, sumR() As Double, sumRh() As Double, sumRl() As Double, dv() As Double, anyD() As Boolean
    ReDim sumW(0 To nb - 1): ReDim sumR(0 To nb - 1): ReDim sumRh(0 To nb - 1): ReDim sumRl(0 To nb - 1)
    ReDim dv(0 To nb - 1): ReDim anyD(0 To nb - 1)
    Dim sDays() As Long, sC() As Double, sH() As Double, sL() As Double, sV() As Double, ns As Long
    Dim j As Long, k As Long, t As Long, tp As Long, w As Double
    For j = 0 To m - 1
        w = caps(j)
        If w > 0 Then
            Dim el As Double: el = Timer - t0: If el < 0 Then el = el + 86400
            Application.StatusBar = "RRG [" & iNo & "/" & nInd & "] " & nm & "  stock " & (j + 1) & "/" & m & " " & syms(j) & _
                                    "  .  fetched " & okC & " ok / " & failC & " none  .  " & Format(el / 60, "0") & " min"
            DoEvents
            Dim ysym As String: ysym = YahooSymbol(CStr(syms(j)))
            ns = FetchOhlcv(ysym, sDays, sC, sH, sL, sV)
            ' TW: a bare or wrongly suffixed code - try the other exchange
            If ns < 2 And capIsShares Then
                If Right(ysym, 4) = "-TWO" Or Right(ysym, 4) = ".TWO" Then
                    ns = FetchOhlcv(Left(ysym, Len(ysym) - 4) & ".TW", sDays, sC, sH, sL, sV)
                ElseIf Right(ysym, 3) = ".TW" Then
                    ns = FetchOhlcv(ysym & "O", sDays, sC, sH, sL, sV)
                Else
                    ns = FetchOhlcv(ysym & ".TW", sDays, sC, sH, sL, sV)
                    If ns < 2 Then ns = FetchOhlcv(ysym & ".TWO", sDays, sC, sH, sL, sV)
                End If
            End If
            If ns >= 2 Then
                If capIsShares Then w = w * sC(ns - 1)          ' shares x last close = market cap
            End If
            If ns >= 2 And capIsShares And w < TW_MIN_CAP Then
                nExcl = nExcl + 1                               ' below the size floor: no weight, no volume
            ElseIf ns >= 2 Then
                nOk = nOk + 1
                incCode(nInc) = Replace(Replace(UCase(CStr(syms(j))), ".TWO", ""), ".TW", ""): incCap(nInc) = w: nInc = nInc + 1
                If bIdx.Exists(sDays(0)) Then
                    t = bIdx(sDays(0)): anyD(t) = True: dv(t) = dv(t) + sC(0) * sV(0)
                End If
                For k = 1 To ns - 1
                    If bIdx.Exists(sDays(k)) And bIdx.Exists(sDays(k - 1)) Then
                        If sC(k - 1) > 0 Then
                            t = bIdx(sDays(k))
                            sumW(t) = sumW(t) + w
                            sumR(t) = sumR(t) + w * (sC(k) / sC(k - 1) - 1)
                            sumRh(t) = sumRh(t) + w * (sH(k) / sC(k - 1) - 1)
                            sumRl(t) = sumRl(t) + w * (sL(k) / sC(k - 1) - 1)
                            dv(t) = dv(t) + sC(k) * sV(k)
                            anyD(t) = True
                        End If
                    End If
                Next k
            Else
                nFail = nFail + 1
            End If
        Else
            nFail = nFail + 1
        End If
    Next j

    ' chain-link
    ReDim days(0 To nb - 1): ReDim c(0 To nb - 1): ReDim h(0 To nb - 1): ReDim l(0 To nb - 1): ReDim v(0 To nb - 1)
    Dim idx As Double: idx = 100
    Dim started As Boolean
    For t = 0 To nb - 1
        If Not started Then
            If anyD(t) Then
                started = True
                days(npts) = bDays(t): c(npts) = idx: h(npts) = idx: l(npts) = idx: v(npts) = dv(t)
                npts = npts + 1
            End If
        ElseIf sumW(t) > 0 Then
            Dim nx As Double: nx = idx * (1 + sumR(t) / sumW(t))
            Dim hx As Double: hx = idx * (1 + sumRh(t) / sumW(t))
            Dim lx As Double: lx = idx * (1 + sumRl(t) / sumW(t))
            If hx < nx Then hx = nx
            If lx > nx Then lx = nx
            days(npts) = bDays(t): c(npts) = nx: h(npts) = hx: l(npts) = lx: v(npts) = dv(t)
            idx = nx
            npts = npts + 1
        End If
    Next t
    If capIsShares Then
        ' column B: the four largest counted members (+N)
        Dim a As Long, b As Long, tc As String, tw As Double, lbl As String
        For a = 0 To nInc - 2
            For b = a + 1 To nInc - 1
                If incCap(b) > incCap(a) Then
                    tc = incCode(a): incCode(a) = incCode(b): incCode(b) = tc
                    tw = incCap(a): incCap(a) = incCap(b): incCap(b) = tw
                End If
            Next b
        Next a
        For a = 0 To nInc - 1
            If a >= 4 Then Exit For
            lbl = lbl & IIf(lbl = "", "", " ") & incCode(a)
        Next a
        If nInc > 4 Then lbl = lbl & " +" & (nInc - 4)
        If lbl = "" Then lbl = "-"
        nm = lbl
    End If
    okC = okC + nOk: failC = failC + nFail: okThis = nOk
    gExclC = gExclC + nExcl
    If npts > 0 Or capIsShares Then Call CacheWrite(ckey, nm, asOf, days, c, h, l, v, npts, nOk, nFail, nExcl)
    IndustryComposite = npts
End Function

' ---- IndustryPx cache: hidden sheet, one row per industry ----
' A code, B as-of day, C ok, D fail, E npts, then five PX_MAXN-wide blocks
' (days, close, high, low, volume) from column PX_FIRST.  Column-blocked so
' one Range read / write moves the whole row.
Private Function CacheSheet(ByVal create As Boolean) As Worksheet
    On Error Resume Next
    Set CacheSheet = ThisWorkbook.Sheets(IND_PX)
    On Error GoTo 0
    If CacheSheet Is Nothing And create Then
        Set CacheSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        CacheSheet.Name = IND_PX
        CacheSheet.cells(1, 1).Value = "code": CacheSheet.cells(1, 2).Value = "asof"
        CacheSheet.cells(1, 3).Value = "ok": CacheSheet.cells(1, 4).Value = "fail": CacheSheet.cells(1, 5).Value = "npts"
        CacheSheet.cells(1, PX_FIRST).Value = "days / close / high / low / volume, " & PX_MAXN & " columns each"
        CacheSheet.Visible = xlSheetHidden
    End If
End Function

Private Function CacheRow(wc As Worksheet, ByVal code As String) As Long
    Dim lastR As Long: lastR = wc.cells(wc.Rows.count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastR
        If StrComp(CStr(wc.cells(r, 1).Value), code, vbTextCompare) = 0 Then CacheRow = r: Exit Function
    Next r
End Function

Private Function CacheRead(ByVal code As String, ByVal asOf As Long, ByRef days() As Long, ByRef c() As Double, _
                           ByRef h() As Double, ByRef l() As Double, ByRef v() As Double, ByRef nOk As Long, ByRef nFail As Long, _
                           ByRef nExcl As Long, ByRef lbl As String) As Long
    Dim wc As Worksheet: Set wc = CacheSheet(False)
    If wc Is Nothing Then Exit Function
    Dim r As Long: r = CacheRow(wc, code)
    If r = 0 Then Exit Function
    If CLng(Val(wc.cells(r, 2).Value)) <> asOf Then Exit Function
    Dim npts As Long: npts = CLng(Val(wc.cells(r, 5).Value))
    If npts < 0 Or npts > PX_MAXN Then Exit Function
    nOk = CLng(Val(wc.cells(r, 3).Value)): nFail = CLng(Val(wc.cells(r, 4).Value))
    nExcl = CLng(Val(wc.cells(r, PX_FIRST + 5 * PX_MAXN + 1).Value))
    lbl = CStr(wc.cells(r, PX_FIRST + 5 * PX_MAXN).Value)
    If npts = 0 Then CacheRead = -1: Exit Function      ' cached empty composite (every member below the floor)
    Dim blk As Variant: blk = wc.Range(wc.cells(r, PX_FIRST), wc.cells(r, PX_FIRST + 5 * PX_MAXN - 1)).Value
    ReDim days(0 To npts - 1): ReDim c(0 To npts - 1): ReDim h(0 To npts - 1): ReDim l(0 To npts - 1): ReDim v(0 To npts - 1)
    Dim k As Long
    For k = 0 To npts - 1
        days(k) = CLng(blk(1, 1 + k))
        c(k) = CDbl(blk(1, 1 + PX_MAXN + k))
        h(k) = CDbl(blk(1, 1 + 2 * PX_MAXN + k))
        l(k) = CDbl(blk(1, 1 + 3 * PX_MAXN + k))
        v(k) = CDbl(blk(1, 1 + 4 * PX_MAXN + k))
    Next k
    CacheRead = npts
End Function

Private Sub CacheWrite(ByVal code As String, ByVal nm As String, ByVal asOf As Long, ByRef days() As Long, ByRef c() As Double, _
                       ByRef h() As Double, ByRef l() As Double, ByRef v() As Double, ByVal npts As Long, ByVal nOk As Long, ByVal nFail As Long, _
                       Optional ByVal nExcl As Long = 0)
    Dim wc As Worksheet: Set wc = CacheSheet(True)
    Dim r As Long: r = CacheRow(wc, code)
    If r = 0 Then r = wc.cells(wc.Rows.count, 1).End(xlUp).Row + 1
    If npts > PX_MAXN Then npts = PX_MAXN
    Dim blk() As Variant: ReDim blk(1 To 1, 1 To 5 * PX_MAXN)
    Dim k As Long
    For k = 0 To npts - 1
        blk(1, 1 + k) = days(k)
        blk(1, 1 + PX_MAXN + k) = c(k)
        blk(1, 1 + 2 * PX_MAXN + k) = h(k)
        blk(1, 1 + 3 * PX_MAXN + k) = l(k)
        blk(1, 1 + 4 * PX_MAXN + k) = v(k)
    Next k
    wc.cells(r, 1).Value = code: wc.cells(r, 2).Value = asOf
    wc.cells(r, 3).Value = nOk: wc.cells(r, 4).Value = nFail: wc.cells(r, 5).Value = npts
    wc.Range(wc.cells(r, PX_FIRST), wc.cells(r, PX_FIRST + 5 * PX_MAXN - 1)).Value = blk
    wc.cells(r, PX_FIRST + 5 * PX_MAXN).Value = nm
    wc.cells(r, PX_FIRST + 5 * PX_MAXN + 1).Value = nExcl
End Sub

' ---- IMAP: import projects/moomoo-plate-list/moomoo_us_plate_stocks.csv ----
' Rebuilds the IndustryMap sheet (A plateCode, B plateName, C symbol,
' D rank, E marketCap, F plateType) from the CSV.  csvPath "" opens a file
' picker, starting at the repo copy when it exists.  The CSV is UTF-8 with
' BOM and Chinese stock names, so it is read through ADODB.Stream; only
' the ASCII columns are kept.
Public Sub ImportIndustryMap(Optional ByVal csvPath As String = "")
    Dim defPath As String
    defPath = Environ("USERPROFILE") & "\OneDrive\" & ChrW(&H684C) & ChrW(&H9762) & "\Claudecode\projects\moomoo-plate-list\moomoo_us_plate_stocks.csv"
    If csvPath = "" Then
        If Dir(defPath) <> "" Then ChDir Left(defPath, InStrRev(defPath, "\") - 1)
        Dim pick As Variant
        pick = Application.GetOpenFilename("moomoo plate stocks CSV (*.csv),*.csv", , "IndustryMap: pick moomoo_us_plate_stocks.csv")
        If VarType(pick) = vbBoolean Then Exit Sub
        csvPath = CStr(pick)
    End If
    If Dir(csvPath) = "" Then Call NavNotify("IMAP: file not found - " & csvPath, True): Exit Sub

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.LoadFromFile csvPath
    Dim txt As String: txt = stm.ReadText
    stm.Close
    If Left(txt, 1) = ChrW(&HFEFF) Then txt = Mid(txt, 2)
    txt = Replace(txt, vbCrLf, vbLf)
    Dim lines() As String: lines = Split(txt, vbLf)
    If UBound(lines) < 1 Then Call NavNotify("IMAP: empty file", True): Exit Sub

    Dim hdr() As String: hdr = SplitCsv(lines(0))
    Dim cCode As Long, cName As Long, cSym As Long, cRank As Long, cCap As Long, cType As Long
    cCode = -1: cName = -1: cSym = -1: cRank = -1: cCap = -1: cType = -1
    Dim j As Long
    For j = 0 To UBound(hdr)
        Select Case hdr(j)
            Case "plateCode": cCode = j
            Case "plateName": cName = j
            Case "symbol": cSym = j
            Case "rank": cRank = j
            Case "marketCapNum": cCap = j
            Case "plateType": cType = j
        End Select
    Next j
    If cCode < 0 Or cSym < 0 Or cCap < 0 Then Call NavNotify("IMAP: header lacks plateCode / symbol / marketCapNum", True): Exit Sub

    Dim out() As Variant: ReDim out(1 To UBound(lines), 1 To 6)
    Dim n As Long, i As Long
    For i = 1 To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim f() As String: f = SplitCsv(lines(i))
            If UBound(f) >= cCap Then
                n = n + 1
                out(n, 1) = f(cCode)
                out(n, 2) = IIf(cName >= 0, f(cName), "")
                out(n, 3) = f(cSym)
                out(n, 4) = IIf(cRank >= 0, Val(f(cRank)), 0)
                out(n, 5) = IIf(IsNumeric(f(cCap)) And f(cCap) <> "", CDbl(f(cCap)), 0)
                out(n, 6) = IIf(cType >= 0, f(cType), "")
            End If
        End If
    Next i

    Dim wm As Worksheet
    On Error Resume Next
    Set wm = ThisWorkbook.Sheets(IND_MAP)
    On Error GoTo 0
    If wm Is Nothing Then
        Set wm = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wm.Name = IND_MAP
    End If
    wm.cells.Clear
    wm.Range("A1:F1").Value = Array("plateCode", "plateName", "symbol", "rank", "marketCap", "plateType")
    wm.Range("A1:F1").Font.Bold = True
    If n > 0 Then wm.Range(wm.cells(2, 1), wm.cells(n + 1, 6)).Value = out
    wm.Columns(2).ColumnWidth = 34: wm.Columns(5).NumberFormat = "#,##0"
    wm.cells(1, 8).Value = "imported " & Format(Now, "yyyy/mm/dd hh:mm") & " from " & csvPath
    ' a new universe invalidates the composites
    Dim wc As Worksheet: Set wc = CacheSheet(False)
    If Not wc Is Nothing Then wc.cells.Clear: wc.cells(1, 1).Value = "code"
    Call NavNotify("IMAP: " & n & " constituents imported into " & IND_MAP & " - run RGI! to build the industry RRG")
End Sub

' RFC-4180 style split: quoted fields may hold commas and doubled quotes.
Private Function SplitCsv(ByVal line As String) As String()
    Dim out() As String: ReDim out(0 To 0)
    Dim n As Long, cur As String, q As Boolean, i As Long, ch As String
    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If q Then
            If ch = """" Then
                If Mid(line, i + 1, 1) = """" Then cur = cur & """": i = i + 1 Else q = False
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" Then
            q = True
        ElseIf ch = "," Then
            ReDim Preserve out(0 To n): out(n) = cur: n = n + 1: cur = ""
        Else
            cur = cur & ch
        End If
    Next i
    ReDim Preserve out(0 To n): out(n) = cur
    SplitCsv = out
End Function

' ================================================================
'  TW GROUPS UNIVERSE  (TG page)
' ================================================================

' The Market = TW groups of tblGroups (Sanner.GetGroupNames / GetSectorTickers),
' one entry per group: tickers(i) = group name (series / focus key), labels(i)
' = the first few member codes, indSyms(i) = members with their exchange
' suffix resolved from the shares tables, indCaps(i) = shares outstanding
' (IndustryComposite multiplies by the last close).  Returns the group count.
Private Function TwGroupUniverse(ByRef tickers() As String, ByRef labels() As String, ByRef groups() As String, _
                                 ByRef indSyms As Variant, ByRef indCaps As Variant, ByRef nStocks As Long) As Long
    Dim gNames As Variant: gNames = Sanner.GetGroupNames("TW")
    If IsEmpty(gNames) Then Exit Function
    Dim n As Long: n = UBound(gNames) + 1
    ReDim tickers(0 To n - 1): ReDim labels(0 To n - 1): ReDim groups(0 To n - 1)
    Dim syms() As Variant, caps() As Variant
    ReDim syms(0 To n - 1): ReDim caps(0 To n - 1)
    Application.StatusBar = "RRG-TW: fetching shares outstanding (TWSE + TPEx)"
    DoEvents
    Dim shares As Object, mkt As Object
    Call TwSharesMap(shares, mkt)
    Dim i As Long, j As Long
    Dim used As Object: Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare
    For i = 0 To n - 1
        ' display / series key = the normalised name; tblGroups keeps its own
        Dim nrm As String: nrm = NormalizeGroupName(CStr(gNames(i)))
        If used.Exists(nrm) Then
            used(nrm) = used(nrm) + 1
            tickers(i) = nrm & " (" & used(nrm) & ")"
        Else
            used(nrm) = 1
            tickers(i) = nrm
        End If
        groups(i) = ""
        Dim tks As Variant: tks = Sanner.GetSectorTickers("TW", CStr(gNames(i)))
        Dim m As Long: m = 0
        If Not IsEmpty(tks) Then m = UBound(tks) + 1
        Dim sArr() As String, cArr() As Double
        ReDim sArr(0 To IIf(m > 0, m - 1, 0)): ReDim cArr(0 To IIf(m > 0, m - 1, 0))
        Dim lbl As String: lbl = ""
        For j = 0 To m - 1
            Dim code As String: code = UCase(Trim(CStr(tks(j))))
            code = Replace(Replace(code, ".TWO", ""), ".TW", "")
            If mkt.Exists(code) Then
                sArr(j) = code & "." & mkt(code)
            Else
                sArr(j) = code                       ' unknown: composite tries .TW then .TWO
            End If
            cArr(j) = IIf(shares.Exists(code), CDbl(shares(code)), 0)
            If j < 4 Then lbl = lbl & IIf(lbl = "", "", " ") & code
            nStocks = nStocks + 1
        Next j
        If m > 4 Then lbl = lbl & " +" & (m - 4)
        labels(i) = lbl
        syms(i) = sArr: caps(i) = cArr
    Next i
    indSyms = syms: indCaps = caps
    TwGroupUniverse = n
End Function

' Shares outstanding for every listed / OTC Taiwan stock:
'   shares(code) = issued shares,  mkt(code) = "TW" (TWSE) / "TWO" (TPEx)
' TWSE: openapi t187ap03_L (company basics; the last field of each record
'   is the issued share count).  TPEx: the daily close table of the last
'   trading day (field 15 = issued shares) - the TPEx openapi company table
'   truncates at random (177 of ~800 records, 2026-09-13), so it is not used.
Private Sub TwSharesMap(ByRef shares As Object, ByRef mkt As Object)
    Set shares = CreateObject("Scripting.Dictionary")
    Set mkt = CreateObject("Scripting.Dictionary")
    Dim http As Object, resp As String, recs() As String, i As Long, f() As String, code As String, v As String
    ' ---- TWSE ----
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://openapi.twse.com.tw/v1/opendata/t187ap03_L", False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    If http.Status = 200 Then resp = http.responseText Else resp = ""
    On Error GoTo 0
    If resp <> "" Then
        resp = Replace(Replace(resp, vbCr, ""), vbLf, "")   ' records are separated by "}," + newline + "{"
        recs = Split(resp, "},{")
        For i = 0 To UBound(recs)
            f = Split(recs(i), """,""")                 ' key":"value pieces
            If UBound(f) >= 2 Then
                code = JsonPieceValue(f(1))              ' 2nd field = company code
                v = JsonPieceValue(f(UBound(f)))         ' last field = issued shares
                If code <> "" And IsNumeric(v) Then
                    shares(code) = CDbl(v): mkt(code) = "TW"
                End If
            End If
        Next i
    End If
    ' ---- TPEx: last trading day's close table ----
    Dim d As Long
    For d = 0 To 7
        resp = ""
        On Error Resume Next
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", "https://www.tpex.org.tw/www/zh-tw/afterTrading/otc?date=" & Format(Date - d, "yyyy/mm/dd") & "&type=EW&response=json", False
        http.setRequestHeader "User-Agent", "Mozilla/5.0"
        http.send
        If http.Status = 200 Then resp = http.responseText
        On Error GoTo 0
        If InStr(resp, """totalCount"":0") = 0 And InStr(resp, "],[") > 0 Then Exit For
        resp = ""
    Next d
    If resp <> "" Then
        recs = Split(resp, "],[")
        For i = 0 To UBound(recs)
            f = Split(recs(i), """,""")
            If UBound(f) >= 14 Then
                code = Replace(Replace(f(0), "[", ""), """", "")
                code = Trim(Mid(code, InStrRev(code, ":") + 1))
                v = Replace(Trim(f(14)), ",", "")
                If IsNumeric(v) And code <> "" And Not mkt.Exists(code) Then
                    shares(code) = CDbl(v): mkt(code) = "TWO"
                End If
            End If
        Next i
    End If
End Sub

' value part of a  key":"value  piece (quotes, braces and brackets stripped)
Private Function JsonPieceValue(ByVal piece As String) As String
    Dim p As Long: p = InStrRev(piece, """:""")
    If p > 0 Then piece = Mid(piece, p + 3)
    piece = Replace(Replace(Replace(Replace(piece, """", ""), "}", ""), "]", ""), ",", "")
    JsonPieceValue = Trim(piece)
End Function

' tblGroups names are a mix of styles ("02. <cjk> - IC<cjk> (IC Design)",
' "<cjk>", "<cjk> Air transportation" ...).  For the TG page only,
' normalise for display (the table itself is untouched):
'   1. drop a leading "NN. " number          2. drop any (parenthetical)
'   3. "CJK category - sub" -> sub            4. drop a trailing U+65CF U+7FA4 ("group")
'   5. "CJK + English translation" -> CJK     6. "CJK + CJK synonym" -> first
Private Function NormalizeGroupName(ByVal nm As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    nm = Replace(nm, ChrW(&H3000), " ")
    re.Pattern = "^\s*\d+\s*[.\uFF0E]\s*": nm = re.Replace(nm, "")
    re.Pattern = "\s*[(\uFF08][^)\uFF09]*[)\uFF09]": nm = re.Replace(nm, "")
    re.Pattern = "\s+": nm = Trim(re.Replace(nm, " "))
    re.Pattern = "^([^\x00-\x7F]+)\s*-\s*(.+)$": nm = re.Replace(nm, "$2")
    nm = Trim(nm)
    Dim tail As String: tail = ChrW(&H65CF) & ChrW(&H7FA4)          ' "group" suffix
    If Len(nm) > 2 And Right(nm, 2) = tail Then nm = Trim(Left(nm, Len(nm) - 2))
    re.Pattern = "^([^\x00-\x7F]+)\s+[A-Za-z]+(\s+[A-Za-z]+)+$": nm = re.Replace(nm, "$1")
    re.Pattern = "^([^\x00-\x7F]+)\s+([^\x00-\x7F]+)$": nm = re.Replace(nm, "$1")
    NormalizeGroupName = Trim(nm)
End Function

' Chart text font of a page (titles, axes, point labels): the TW groups page
' carries Chinese names, which Consolas has no glyphs for (Excel falls back
' per character and the labels look ragged), so its charts use Noto Sans TC.
' The table cells and the corner / quadrant captions stay Consolas everywhere.
Private Function PageFont(ws As Worksheet) As String
    If ws.Name = TWG_SHEET Then PageFont = "Noto Sans TC" Else PageFont = "Consolas"
End Function

' First column of the tail data block on a page (set by BuildRRGCore so the
' block clears the charts, which are wider on the industry / TW pages);
' AA when the page predates the mark.  DATE and SEQ sit just left of it.
Private Function DataCol(ws As Worksheet) As Long
    DataCol = CLng(Val(GetMark(ws, DATA_COL_MARK)))
    If DataCol < DATA_COL_MIN Then DataCol = DATA_COL_MIN
    DataCol = DataCol + NavLeft(ws)                  ' the mark is the build-time column; NavAdd inserts column A after
End Function

Private Function DateCol(ws As Worksheet) As Long
    DateCol = DataCol(ws) - 1
End Function

Private Function SeqCol(ws As Worksheet) As Long
    SeqCol = DataCol(ws) - 2
End Function
