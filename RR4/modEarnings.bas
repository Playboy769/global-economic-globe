Attribute VB_Name = "modEarnings"
Option Explicit

' ================================================================
'  EARNINGS (nav code E, E! = refetch) - sheet "Earnings"   2026-09-13
' ----------------------------------------------------------------
'  One ticker at a time.  Page layout (page rows; the nav bar adds a blank
'  row + 3 bar rows above and a blank column A - read inputs through
'  NavOffset / NavLeft):
'    row 1   EARNINGS title | context (entity, market, build time)
'    row 2   TICKER <GO> label       MKT label        (label ABOVE input)
'    row 3   ticker input (A)        market input (C: blank/AUTO, US, TW)
'    row 5-7 EARNINGS DATES: next earnings date, days to go, status
'            (CONFIRMED / ESTIMATE from Yahoo calendarEvents, PROJECTED when
'            Yahoo is unavailable), last earnings call, last filing
'    row 9-  QUARTERLY SNAPSHOT: the last 12 quarters, oldest -> newest, with
'            period end / filed / form rows and the same sections, metrics and
'            formulas as SEC-Filing-Fetcher's QuarterlySnapshot sheet
'            (modCharts.BuildSnapshotTableInto, quarterly branch; the 3Y/5Y
'            CAGR rows are left out - they are always N/A on quarters), plus
'            the rows the old Company research deep-dive showed (gross profit /
'            margin, operating & net margin, total liabilities, interest
'            expense / coverage, CFO).  P/E and P/S range band to the right.
'    TW only: monthly revenue (24 months) under the table.
'
'  Data: US = shared-vba/modSECData (SEC companyfacts), TW = shared-vba/
'  modMOPSData (MOPS t164sb01 iXBRL + t21sc03), prices = modPrices (Yahoo).
'  Q4: SEC never tags a standalone Q4 duration, so a Q4 column is the 10-K
'  full year minus the Q3 year-to-date (flows) / the 10-K's year-end value
'  (balances), shown in italics.  TW Q4 gets the same treatment only when
'  MOPS has no standalone Q4 revenue.
'  Two deliberate differences from SEC-Filing-Fetcher (both are bugs there):
'  10-Q cash-flow tags are year-to-date and are de-cumulated to quarters here
'  (see UsVal), and EPS / DPS / share counts from filings made before a stock
'  split are restated with Yahoo's split history, so they match the
'  split-adjusted Yahoo prices (NVDA 10:1, 2024-06-10).
'  TW "Filed" is the statutory deadline (5/15, 8/14, 11/14, 3/31), not the
'  actual filing date - MOPS t164sb01 does not expose it.
'
'  This replaces the lower half of "Company research": typing a ticker there
'  (or double-clicking one in its scan table) now lands here.
'  Pure ASCII (VBE import rule).
' ================================================================

Public Const EARN_SHEET As String = "Earnings"

Private Const PG_TITLE As Long = 1
Private Const PG_LBL As Long = 2
Private Const PG_IN As Long = 3
Private Const COL_TK As Long = 1                 ' A at draw time (B once the bar adds column A)
Private Const COL_MKT As Long = 3                ' C
Private Const PG_DT_TITLE As Long = 5
Private Const PG_DT_LBL As Long = 6
Private Const PG_DT_VAL As Long = 7
Private Const PG_TBL_TITLE As Long = 9
Private Const PG_TBL_HDR As Long = 10

Private Const N_QTR As Long = 12
Private Const TW_MONTHS As Long = 24
Private Const LAST_CLEAR_COL As Long = 40

Private Const UA As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
Private Const FONT_FACE As String = "Consolas"
Private Const CLR_TEXT As Long = 14540253        ' RGB(221,221,221)
Private Const CLR_MUTED As Long = 8553090        ' RGB(130,130,130)
Private Const CLR_BANNER As Long = 1842204       ' RGB(28,28,28)
Private Const CLR_FLAG As Long = 49407           ' RGB(255,192,0)

' the loaded quarters, 1 = oldest
Private m_n As Long
Private m_lbl() As String
Private m_end() As String                        ' yyyy-mm-dd
Private m_filed() As String                      ' yyyy-mm-dd
Private m_form() As String
Private m_q4() As Boolean                        ' derived Q4 column
Private m_v As Object                            ' metric key -> Variant(1..m_n), "" = missing
Private m_splits As Collection                   ' Array(split date, ratio) from Yahoo

' ----------------------------------------------------------------
'  Entry points
' ----------------------------------------------------------------
Public Sub ShowEarnings(ByVal rawTicker As String, Optional ByVal marketOverride As String = "", Optional ByVal goPage As Boolean = False)
    Dim ws As Worksheet: Set ws = EnsureEarningsSheet()
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fail

    Call NavStrip(ws)
    Dim tk As String: tk = UCase$(Trim$(rawTicker))
    Dim mktIn As String: mktIn = UCase$(Trim$(marketOverride))
    Call DrawShell(ws, tk, mktIn)
    If tk = "" Then GoTo Done

    Dim mkt As String
    If Right$(tk, 4) = ".TWO" Then
        tk = Left$(tk, Len(tk) - 4): mkt = "TW"
    ElseIf Right$(tk, 3) = ".TW" Then
        tk = Left$(tk, Len(tk) - 3): mkt = "TW"
    Else
        mkt = DetectMarket(tk, mktIn)
    End If

    Call NavNotify("EARNINGS fetching " & tk & " (" & mkt & ") ...")
    Dim entity As String, srcNote As String, errMsg As String, ok As Boolean
    If mkt = "TW" Then
        ok = LoadTw(tk, entity, srcNote, errMsg)
    Else
        ok = LoadUs(tk, entity, srcNote, errMsg)
    End If
    If Not ok Then
        Call WriteFlag(ws, PG_TBL_TITLE, "No data for " & tk & ": " & errMsg)
        Call NavNotify("EARNINGS " & tk & ": " & errMsg, True)
        GoTo Done
    End If

    Dim sym As String
    Dim prices As Object: Set prices = LoadPrices(tk, mkt, sym)
    With ws.cells(PG_TITLE, 3)
        .Value = entity & "  (" & tk & ")  .  " & mkt & "  .  built " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = CLR_MUTED
    End With
    Call RenderDates(ws, mkt, IIf(sym = "", tk, sym))
    Dim lastRow As Long
    lastRow = RenderTable(ws, mkt, prices, srcNote)
    If mkt = "TW" Then lastRow = RenderMonthly(ws, lastRow + 2, tk)
    Call NavNotify("EARNINGS " & tk & " done - " & m_n & " quarters")

Done:
    Call FinishPage(ws)
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
    If goPage Then Call NavGoto("E", ws)
    Exit Sub
Fail:
    Dim msg As String: msg = Err.Description
    Resume FailOut                                   ' leave the handler before doing anything else
FailOut:
    On Error Resume Next
    Call WriteFlag(ws, PG_TBL_TITLE, "Error: " & msg)
    Call NavNotify("EARNINGS error: " & msg, True)
    GoTo Done
End Sub

' E! - refetch whatever the page's inputs hold (or just build the empty page).
Public Sub RefreshEarnings()
    Dim ws As Worksheet: Set ws = EnsureEarningsSheet()
    Dim r As Long: r = PG_IN + NavOffset(ws)
    Dim tk As String, mk As String
    tk = CStr(ws.cells(r, COL_TK + NavLeft(ws)).Value)
    mk = CStr(ws.cells(r, COL_MKT + NavLeft(ws)).Value)
    Call ShowEarnings(tk, mk, True)
End Sub

' Sheet event (Worksheet_Change, written by EnsureSheetCode).
Public Sub EarningsChange(ws As Worksheet, ByVal Target As Range)
    If Target.CountLarge > 4 Then Exit Sub
    Dim r As Long: r = PG_IN + NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim inputs As Range
    Set inputs = Union(ws.cells(r, COL_TK + lc), ws.cells(r, COL_MKT + lc))
    If Intersect(Target, inputs) Is Nothing Then Exit Sub
    Call ShowEarnings(CStr(ws.cells(r, COL_TK + lc).Value), CStr(ws.cells(r, COL_MKT + lc).Value))
End Sub

' ----------------------------------------------------------------
'  Sheet / shell
' ----------------------------------------------------------------
Private Function EnsureEarningsSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(EARN_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Dim prev As Object: Set prev = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = EARN_SHEET
        ActiveWindow.DisplayGridlines = False
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If
    Set EnsureEarningsSheet = ws
End Function

Private Sub DrawShell(ws As Worksheet, ByVal tk As String, ByVal mkt As String)
    ws.cells.Clear
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = FONT_FACE
        .Font.Size = 9
        .Font.Color = CLR_TEXT
        .VerticalAlignment = xlCenter
        .RowHeight = 16
    End With
    With ws.cells(PG_TITLE, 1)
        .Value = "EARNINGS"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 14
    End With
    ws.Rows(PG_TITLE).RowHeight = 24

    Call StackLabel(ws.cells(PG_LBL, COL_TK), "TICKER <GO>")
    Call StackLabel(ws.cells(PG_LBL, COL_MKT), "MKT")
    ws.Rows(PG_LBL).RowHeight = 18
    ws.Rows(PG_IN).RowHeight = 20
    Call InputCell(ws.cells(PG_IN, COL_TK), tk)
    Call InputCell(ws.cells(PG_IN, COL_MKT), mkt)
    With ws.cells(PG_IN, COL_MKT + 2)
        .Value = "US ticker (NVDA) or TW code (2330 / 3374.TWO) + Enter.  MKT: blank = auto, US, TW"
        .Font.Color = CLR_MUTED
    End With

    ws.Columns(1).ColumnWidth = 30
    Dim c As Long
    For c = 2 To LAST_CLEAR_COL: ws.Columns(c).ColumnWidth = 12: Next c
End Sub

Private Sub StackLabel(cell As Range, ByVal txt As String)
    With cell
        .Value = txt
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
End Sub

Private Sub InputCell(cell As Range, ByVal v As String)
    With cell
        .NumberFormat = "@"
        .Value = v
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub FinishPage(ws As Worksheet)
    On Error Resume Next
    Call NavAdd(ws, "E")
    Call EnsureSheetCode(ws)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
'  US loader (SEC XBRL)
' ----------------------------------------------------------------
Private Function LoadUs(ByVal tk As String, ByRef entity As String, ByRef srcNote As String, ByRef errMsg As String) As Boolean
    Dim r As Object
    Set r = modSECData.GetXbrlFinancials(tk, 6, N_QTR + 3)
    If Not r("ok") Then errMsg = CStr(r("error")): Exit Function
    entity = CStr(r("entityName"))
    Dim maps As Object: Set maps = r("maps")
    Dim allMaps As Variant: allMaps = r("allMaps")

    ' ---- columns: every 10-Q + a derived Q4 per fiscal year with Q1-Q3 and a 10-K ----
    Dim cols As New Collection, byFy As Object
    Set byFy = CreateObject("Scripting.Dictionary")
    Dim f As Variant, c As Object, parts As Variant
    For Each f In r("filings10Q")
        Set c = CreateObject("Scripting.Dictionary")
        c("accn") = CStr(f("accn")): c("reportDate") = CStr(f("reportDate"))
        c("form") = CStr(f("form")): c("filed") = CStr(f("filingDate"))
        c("q4") = False
        parts = Split(modSECData.XbrlFyFp(allMaps, c("accn"), c("reportDate"), c("form")) & "|", "|")
        c("fy") = parts(0): c("fp") = parts(1)
        cols.Add c
        If c("fy") <> "" Then
            If Not byFy.Exists(c("fy")) Then Set byFy(c("fy")) = CreateObject("Scripting.Dictionary")
            Set byFy(c("fy"))(c("fp")) = c
        End If
    Next f
    For Each f In r("filings10K")
        parts = Split(modSECData.XbrlFyFp(allMaps, CStr(f("accn")), CStr(f("reportDate")), CStr(f("form"))) & "|", "|")
        If parts(0) <> "" And parts(1) = "FY" Then
            If byFy.Exists(parts(0)) Then
                If byFy(parts(0)).Exists("Q1") And byFy(parts(0)).Exists("Q2") And byFy(parts(0)).Exists("Q3") Then
                    Set c = CreateObject("Scripting.Dictionary")
                    c("accn") = CStr(f("accn")): c("reportDate") = CStr(f("reportDate"))
                    c("form") = CStr(f("form")): c("filed") = CStr(f("filingDate"))
                    c("q4") = True: c("fy") = parts(0): c("fp") = "Q4"
                    Set c("q1") = byFy(parts(0))("Q1")
                    Set c("q2") = byFy(parts(0))("Q2")
                    Set c("q3") = byFy(parts(0))("Q3")
                    cols.Add c
                End If
            End If
        End If
    Next f
    If cols.count = 0 Then errMsg = "no 10-Q filings on record": Exit Function

    Dim arr() As Object, i As Long, j As Long, n As Long
    n = cols.count
    ReDim arr(1 To n)
    For i = 1 To n: Set arr(i) = cols(i): Next i
    Dim key As Object
    For i = 2 To n                                   ' ascending by reportDate (ISO text sorts)
        Set key = arr(i): j = i - 1
        Do While j >= 1
            If CStr(arr(j)("reportDate")) > CStr(key("reportDate")) Then
                Set arr(j + 1) = arr(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = key
    Next i

    Dim first As Long: first = 1
    If n > N_QTR Then first = n - N_QTR + 1
    Call AllocCols(n - first + 1)
    For i = first To n
        j = i - first + 1
        Set c = arr(i)
        If c("fy") <> "" Then m_lbl(j) = "FY" & c("fy") & c("fp") Else m_lbl(j) = Left$(c("reportDate"), 7)
        m_end(j) = c("reportDate"): m_filed(j) = c("filed")
        m_form(j) = IIf(c("q4"), c("form") & " (Q4)", c("form"))
        m_q4(j) = c("q4")
    Next i

    Call LoadSplits(tk)
    Dim flows As Variant, stocks As Variant, k As Variant, a() As Variant
    flows = FlowKeys(): stocks = StockKeys()
    For Each k In flows
        ReDim a(1 To m_n)
        For i = first To n: a(i - first + 1) = UsVal(maps, CStr(k), arr(i), True, byFy): Next i
        m_v(CStr(k)) = a
    Next k
    For Each k In stocks
        ReDim a(1 To m_n)
        For i = first To n: a(i - first + 1) = UsVal(maps, CStr(k), arr(i), False, byFy): Next i
        m_v(CStr(k)) = a
    Next k
    ReDim a(1 To m_n)                                ' a rate cannot be backed out -> blank on Q4
    For i = first To n
        If arr(i)("q4") Then a(i - first + 1) = "" Else a(i - first + 1) = Num(modSECData.XbrlValue(maps, "EffectiveTaxRate", arr(i)("accn"), arr(i)("reportDate"), arr(i)("form")))
    Next i
    m_v("EffectiveTaxRate") = a

    srcNote = "Source: SEC EDGAR companyfacts XBRL (CIK " & CStr(r("cik")) & "), prices Yahoo. " & _
              "Flows are standalone quarters (year-to-date 10-Q cash-flow tags are de-cumulated). " & _
              "Italic Q4 = 10-K full year minus the Q3 year-to-date (balances: 10-K year-end); tax rate blank there. " & _
              "EPS / DPS / shares restated to the post-split basis (Yahoo split history). Ratios on single-quarter figures, not annualised."
    LoadUs = True
End Function

' One column's value for concept k.
'  balances: the fact as of the period end.
'  flows:    the STANDALONE quarter.  A 10-Q's cash-flow statement is tagged
'            year-to-date only (no 3-month duration), and the fact picker then
'            returns the YTD figure - SEC-Filing-Fetcher shows those YTD numbers
'            as quarters and its derived Q4 goes negative (NVDA CFO).  Here a
'            fact spanning > 120 days is treated as YTD and the prior quarters'
'            YTD is subtracted; Q4 = full year minus the Q3 YTD.
'  Per-share values and share counts are put on the post-split basis (SplitAdj).
Private Function UsVal(maps As Object, ByVal k As String, ByVal col As Object, ByVal isFlow As Boolean, byFy As Object) As Variant
    Dim days As Long, v As Variant
    v = UsFact(maps, k, col, days)
    If Not isFlow Then UsVal = v: Exit Function
    If col("q4") Then
        UsVal = Sb(v, UsYtd(maps, k, byFy, CStr(col("fy")), 3))
        Exit Function
    End If
    Dim q As Long: q = QNum(CStr(col("fp")))
    If Not HasVal(v) Then UsVal = "": Exit Function
    If days <= 120 Or q <= 1 Or CStr(col("fy")) = "" Then UsVal = v: Exit Function
    UsVal = Sb(v, UsYtd(maps, k, byFy, CStr(col("fy")), q - 1))
End Function

' Year-to-date total through quarter q of fiscal year fy (0 for q = 0).
Private Function UsYtd(maps As Object, ByVal k As String, byFy As Object, ByVal fy As String, ByVal q As Long) As Variant
    UsYtd = ""
    If q = 0 Then UsYtd = 0: Exit Function
    If fy = "" Then Exit Function
    If Not byFy.Exists(fy) Then Exit Function
    If Not byFy(fy).Exists("Q" & q) Then Exit Function
    Dim days As Long, v As Variant
    v = UsFact(maps, k, byFy(fy)("Q" & q), days)
    If Not HasVal(v) Then Exit Function
    If days > 120 Then
        UsYtd = v
    Else
        UsYtd = Ad(UsYtd(maps, k, byFy, fy, q - 1), v)
    End If
End Function

' Same fact choice as modSECData.LookupConceptValue (end = period end, form
' match preferred, latest start wins), but also reports the fact's duration.
Private Function UsFact(maps As Object, ByVal k As String, ByVal col As Object, ByRef days As Long) As Variant
    UsFact = "": days = 0
    If Not maps.Exists(k) Then Exit Function
    Dim m As Object: Set m = maps(k)
    If m Is Nothing Then Exit Function
    Dim accn As String: accn = CStr(col("accn"))
    If Not m.Exists(accn) Then Exit Function
    Dim facts As Collection: Set facts = m(accn)
    Dim rd As String: rd = CStr(col("reportDate"))
    Dim fm As String: fm = CStr(col("form"))
    Dim best As String, bestStart As String, e As Variant, s As String, pass As Long
    For pass = 1 To 2
        For Each e In facts
            If ExtractJsonString(CStr(e), "end") = rd Then
                If pass = 2 Or ExtractJsonString(CStr(e), "form") = fm Then
                    s = ExtractJsonString(CStr(e), "start")
                    If best = "" Or s > bestStart Then best = CStr(e): bestStart = s
                End If
            End If
        Next e
        If best <> "" Then Exit For
    Next pass
    If best = "" Then best = CStr(facts(facts.count))
    If bestStart <> "" And IsDate(bestStart) And IsDate(rd) Then days = DateDiff("d", CDate(bestStart), CDate(rd))
    UsFact = SplitAdj(k, Num(ExtractJsonValueRaw(best, "val")), CStr(col("filed")))
End Function

Private Function QNum(ByVal fp As String) As Long
    Select Case fp
        Case "Q1": QNum = 1
        Case "Q2": QNum = 2
        Case "Q3": QNum = 3
    End Select
End Function

' ---- stock splits (Yahoo chart events): filings made before a split report
' pre-split per-share values and share counts; Yahoo prices are split-adjusted.
Private Sub LoadSplits(ByVal sym As String)
    Set m_splits = New Collection
    On Error GoTo Fin
    Dim js As String
    js = HttpGet("https://query1.finance.yahoo.com/v8/finance/chart/" & sym & "?range=20y&interval=3mo&events=splits")
    Dim p As Long: p = InStr(js, """splits"":{")
    If p = 0 Then Exit Sub
    Dim blk As String: blk = Mid$(js, p, InStr(p, js, "}}") - p + 2)
    Dim q As Long: q = 1
    Do
        q = InStr(q, blk, """date"":")
        If q = 0 Then Exit Do
        Dim ts As Double: ts = Val(Mid$(blk, q + 7, 12))
        Dim nu As Double, de As Double
        nu = Val(Mid$(blk, InStr(q, blk, """numerator"":") + 12, 12))
        de = Val(Mid$(blk, InStr(q, blk, """denominator"":") + 14, 12))
        If ts > 0 And nu > 0 And de > 0 Then m_splits.Add Array(DateAdd("s", ts, DateSerial(1970, 1, 1)), nu / de)
        q = q + 7
    Loop
Fin:
End Sub

Private Function SplitAdj(ByVal k As String, ByVal v As Variant, ByVal filed As String) As Variant
    SplitAdj = v
    If Not HasVal(v) Then Exit Function
    If m_splits Is Nothing Then Exit Function
    If m_splits.count = 0 Or Not IsDate(filed) Then Exit Function
    Dim f As Double: f = 1
    Dim sp As Variant
    For Each sp In m_splits
        If CDate(filed) < sp(0) Then f = f * sp(1)
    Next sp
    If f = 1 Then Exit Function
    Select Case k
        Case "Eps", "Dividends": SplitAdj = CDbl(v) / f
        Case "Shares":           SplitAdj = CDbl(v) * f
    End Select
End Function

' ----------------------------------------------------------------
'  TW loader (MOPS iXBRL)
' ----------------------------------------------------------------
Private Function LoadTw(ByVal coId As String, ByRef entity As String, ByRef srcNote As String, ByRef errMsg As String) As Boolean
    Dim r As Object
    Set r = modMOPSData.GetTwFinancials(coId, N_QTR + 3)
    If Not r("ok") Then errMsg = CStr(r("error")): Exit Function
    entity = CStr(r("entityName"))
    Dim rows As Collection: Set rows = r("quarterlyRows")          ' newest first
    Dim annual As Collection: Set annual = r("annualRows")
    If rows.count = 0 Then errMsg = "no MOPS quarterly data": Exit Function

    Dim take As Long: take = rows.count
    If take > N_QTR Then take = N_QTR
    Call AllocCols(take)
    Dim spec() As Object, i As Long, j As Long
    ReDim spec(1 To take)
    For i = 1 To take
        j = take - i + 1                                          ' oldest -> column 1
        Set spec(j) = rows(i)
    Next i
    Dim qEnd As Date
    For j = 1 To take
        qEnd = CDate(spec(j)("reportDate"))
        m_lbl(j) = CStr(spec(j)("fyLabel")) & CStr(spec(j)("fpLabel"))
        m_end(j) = Format$(qEnd, "yyyy-mm-dd")
        m_filed(j) = Format$(TwDeadline(qEnd), "yyyy-mm-dd")
        m_form(j) = "MOPS " & CStr(spec(j)("fpLabel"))
        m_q4(j) = False
        If CStr(spec(j)("fpLabel")) = "Q4" Then
            If Not HasVal(TwMetric(spec(j), "Revenue")) Then m_q4(j) = True
        End If
    Next j

    Dim flows As Variant, stocks As Variant, k As Variant, a() As Variant
    flows = FlowKeys(): stocks = StockKeys()
    For Each k In flows
        ReDim a(1 To m_n)
        For j = 1 To m_n
            If m_q4(j) Then
                a(j) = TwQ4Flow(rows, annual, CStr(spec(j)("fyLabel")), CStr(k))
            Else
                a(j) = TwMetric(spec(j), CStr(k))
            End If
        Next j
        m_v(CStr(k)) = a
    Next k
    For Each k In stocks
        ReDim a(1 To m_n)
        For j = 1 To m_n
            a(j) = TwMetric(spec(j), CStr(k))
            If m_q4(j) And Not HasVal(a(j)) Then a(j) = TwMetric(FindRow(annual, CStr(spec(j)("fyLabel")), "FY"), CStr(k))
        Next j
        m_v(CStr(k)) = a
    Next k
    ReDim a(1 To m_n)
    For j = 1 To m_n
        If m_q4(j) Then a(j) = "" Else a(j) = TwMetric(spec(j), "EffectiveTaxRate")
    Next j
    m_v("EffectiveTaxRate") = a

    srcNote = "Source: MOPS t164sb01 inline XBRL (consolidated), prices Yahoo. Filed = statutory deadline, not the actual filing date. " & _
              "Italic Q4 = FY minus Q1-Q3 (only when MOPS has no standalone Q4). Mid-year CFO / CapEx can be blank (MOPS tags them year-to-date). " & _
              "Ratios on single-quarter figures, not annualised."
    LoadTw = True
End Function

Private Function TwMetric(ByVal spec As Object, ByVal k As String) As Variant
    TwMetric = ""
    If spec Is Nothing Then Exit Function
    Dim m As Object: Set m = spec("metrics")
    If m.Exists(k) Then TwMetric = Num(m(k))
End Function

Private Function FindRow(rows As Collection, ByVal fy As String, ByVal fp As String) As Object
    Dim s As Variant
    For Each s In rows
        If CStr(s("fyLabel")) = fy And CStr(s("fpLabel")) = fp Then Set FindRow = s: Exit Function
    Next s
End Function

Private Function TwQ4Flow(rows As Collection, annual As Collection, ByVal fy As String, ByVal k As String) As Variant
    TwQ4Flow = Sb(Sb(Sb(TwMetric(FindRow(annual, fy, "FY"), k), _
                        TwMetric(FindRow(rows, fy, "Q1"), k)), _
                        TwMetric(FindRow(rows, fy, "Q2"), k)), _
                        TwMetric(FindRow(rows, fy, "Q3"), k))
End Function

' Taiwan filing deadlines for a quarter ending on qEnd.
Private Function TwDeadline(ByVal qEnd As Date) As Date
    Select Case Month(qEnd)
        Case 3:  TwDeadline = DateSerial(Year(qEnd), 5, 15)
        Case 6:  TwDeadline = DateSerial(Year(qEnd), 8, 14)
        Case 9:  TwDeadline = DateSerial(Year(qEnd), 11, 14)
        Case Else: TwDeadline = DateSerial(Year(qEnd) + 1, 3, 31)
    End Select
End Function

' ----------------------------------------------------------------
'  Shared data helpers
' ----------------------------------------------------------------
Private Sub AllocCols(ByVal n As Long)
    m_n = n
    ReDim m_lbl(1 To n): ReDim m_end(1 To n): ReDim m_filed(1 To n)
    ReDim m_form(1 To n): ReDim m_q4(1 To n)
    Set m_v = CreateObject("Scripting.Dictionary")
End Sub

Private Function FlowKeys() As Variant
    FlowKeys = Array("Revenue", "GrossProfit", "OperatingIncome", "NetIncome", "Eps", "Cogs", "Da", _
                     "CapEx", "Cfo", "Dividends", "InterestExpense")
End Function

Private Function StockKeys() As Variant
    StockKeys = Array("Shares", "Inventory", "Ar", "AccountsPayable", "CurrentAssets", "CurrentLiabilities", _
                      "LongTermDebt", "ShortTermDebt", "StockholdersEquity", "Cash", "Assets", "Liabilities")
End Function

Private Function LoadPrices(ByVal tk As String, ByVal mkt As String, ByRef sym As String) As Object
    Dim p As Object
    Dim fromD As Date: fromD = CDate(m_end(1)) - 10
    On Error Resume Next
    If mkt = "TW" Then
        sym = tk & ".TW"
        Set p = FetchDailyPrices(sym, fromD, Date)
        If p Is Nothing Then Set p = CreateObject("Scripting.Dictionary")
        If p.count = 0 Then
            sym = tk & ".TWO"
            Set p = FetchDailyPrices(sym, fromD, Date)
        End If
    Else
        sym = tk
        Set p = FetchDailyPrices(sym, fromD, Date)
    End If
    On Error GoTo 0
    If p Is Nothing Then Set p = CreateObject("Scripting.Dictionary")
    If p.count = 0 Then sym = ""
    Set LoadPrices = p
End Function

Private Function DetectMarket(ByVal tk As String, ByVal override As String) As String
    If override = "US" Or override = "TW" Then DetectMarket = override: Exit Function
    Dim i As Long, allDigits As Boolean
    allDigits = (Len(tk) >= 4 And Len(tk) <= 6)
    For i = 1 To Len(tk)
        If InStr("0123456789", Mid$(tk, i, 1)) = 0 Then allDigits = False: Exit For
    Next i
    DetectMarket = IIf(allDigits, "TW", "US")
End Function

' ---- blank-safe arithmetic ("" = missing; VBA And does not short-circuit) ----
Private Function HasVal(ByVal v As Variant) As Boolean
    If IsObject(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Then Exit Function
    End If
    HasVal = IsNumeric(v)
End Function

Private Function Num(ByVal v As Variant) As Variant
    If HasVal(v) Then Num = CDbl(v) Else Num = ""
End Function

Private Function Dv(ByVal a As Variant, ByVal b As Variant) As Variant
    Dv = ""
    If Not HasVal(a) Then Exit Function
    If Not HasVal(b) Then Exit Function
    If CDbl(b) = 0 Then Exit Function
    Dv = CDbl(a) / CDbl(b)
End Function

Private Function Mu(ByVal a As Variant, ByVal b As Variant) As Variant
    Mu = ""
    If HasVal(a) Then If HasVal(b) Then Mu = CDbl(a) * CDbl(b)
End Function

Private Function Ad(ByVal a As Variant, ByVal b As Variant) As Variant
    Ad = ""
    If HasVal(a) Then If HasVal(b) Then Ad = CDbl(a) + CDbl(b)
End Function

Private Function Sb(ByVal a As Variant, ByVal b As Variant) As Variant
    Sb = ""
    If HasVal(a) Then If HasVal(b) Then Sb = CDbl(a) - CDbl(b)
End Function

Private Function Growth(ByVal cur As Variant, ByVal prev As Variant) As Variant
    Dim q As Variant: q = Dv(cur, prev)
    If HasVal(q) Then Growth = q - 1 Else Growth = ""
End Function

Private Function Pos(ByVal v As Variant) As Boolean
    If HasVal(v) Then Pos = (CDbl(v) > 0)
End Function

' ----------------------------------------------------------------
'  Earnings dates
' ----------------------------------------------------------------
Private Sub RenderDates(ws As Worksheet, ByVal mkt As String, ByVal ySym As String)
    Call Section(ws, PG_DT_TITLE, "EARNINGS DATES")
    ws.cells(PG_DT_LBL, 1).Value = "NEXT EARNINGS"
    ws.cells(PG_DT_LBL, 2).Value = "DAYS"
    ws.cells(PG_DT_LBL, 3).Value = "STATUS"
    ws.cells(PG_DT_LBL, 4).Value = "LAST CALL"
    ws.cells(PG_DT_LBL, 5).Value = "LAST FILED"
    ws.cells(PG_DT_LBL, 7).Value = "SOURCE"
    With ws.Range(ws.cells(PG_DT_LBL, 1), ws.cells(PG_DT_LBL, 7)).Font
        .Color = RR4_ACCENT: .Bold = True
    End With

    Dim nd As String, est As Boolean, lastCall As String, status As String, src As String
    If YahooCalendar(ySym, nd, est, lastCall) Then
        status = IIf(est, "ESTIMATE", "CONFIRMED")
        src = "Yahoo calendarEvents (" & ySym & ")"
    Else
        nd = Format$(ProjectNext(mkt), "yyyy-mm-dd")
        status = "PROJECTED"
        If mkt = "TW" Then
            src = "Yahoo unavailable - next statutory filing deadline"
        Else
            src = "Yahoo unavailable - same quarter's filing date last year + 1y (SEC)"
        End If
    End If

    Dim dNext As Date: dNext = CDate(nd)
    With ws.cells(PG_DT_VAL, 1)
        .Value = dNext: .NumberFormat = "yyyy/mm/dd ddd"
        .Font.Bold = True: .Font.Size = 11: .HorizontalAlignment = xlLeft
    End With
    Dim days As Long: days = CLng(dNext - Date)
    With ws.cells(PG_DT_VAL, 2)
        .Value = days: .NumberFormat = "0"
        .Font.Bold = True: .HorizontalAlignment = xlLeft
        If days >= 0 And days <= 14 Then .Font.Color = RR4_ACCENT
    End With
    With ws.cells(PG_DT_VAL, 3)
        .Value = status: .Font.Bold = True
        If status = "CONFIRMED" Then .Font.Color = RGB(80, 200, 120) Else .Font.Color = RGB(249, 168, 37)
    End With
    If lastCall <> "" Then
        ws.cells(PG_DT_VAL, 4).Value = CDate(lastCall)
        ws.cells(PG_DT_VAL, 4).NumberFormat = "yyyy/mm/dd"
        ws.cells(PG_DT_VAL, 4).HorizontalAlignment = xlLeft
    Else
        ws.cells(PG_DT_VAL, 4).Value = "-"
    End If
    ws.cells(PG_DT_VAL, 5).Value = m_filed(m_n) & "  " & m_form(m_n)
    With ws.cells(PG_DT_VAL, 7)
        .Value = src: .Font.Color = CLR_MUTED
    End With
    ws.Rows(PG_DT_VAL).RowHeight = 20
End Sub

Private Function ProjectNext(ByVal mkt As String) As Date
    Dim d As Date
    If mkt = "TW" Then
        Dim qe As Date: qe = CDate(m_end(m_n))
        Do
            qe = DateSerial(Year(qe), Month(qe) + 4, 0)          ' last day of the quarter after
            d = TwDeadline(qe)
        Loop While d < Date
    Else
        If m_n >= 4 Then
            d = CDate(m_filed(m_n - 3)) + 364
        Else
            d = CDate(m_filed(m_n)) + 91
        End If
        Do While d < Date: d = d + 91: Loop
    End If
    ProjectNext = d
End Function

' Yahoo quoteSummary calendarEvents needs a session cookie + crumb:
' fc.yahoo.com hands out the cookie, /v1/test/getcrumb turns it into a crumb.
Private Function YahooCalendar(ByVal sym As String, ByRef nd As String, ByRef est As Boolean, ByRef lastCall As String) As Boolean
    On Error GoTo Bad
    Dim st As Long, hdrs As String, body As String
    body = YGet("https://fc.yahoo.com/", "", st, hdrs)
    Dim cookie As String: cookie = CookiesFrom(hdrs)
    If cookie = "" Then Exit Function
    Dim crumb As String
    crumb = Trim$(YGet("https://query2.finance.yahoo.com/v1/test/getcrumb", cookie, st, hdrs))
    If st <> 200 Or crumb = "" Or Len(crumb) > 40 Or InStr(crumb, "{") > 0 Then Exit Function
    body = YGet("https://query2.finance.yahoo.com/v10/finance/quoteSummary/" & sym & _
                "?modules=calendarEvents&crumb=" & UrlEnc(crumb), cookie, st, hdrs)
    If st <> 200 Then Exit Function
    nd = FmtAfter(body, """earningsDate""")
    lastCall = FmtAfter(body, """earningsCallDate""")
    est = (InStr(body, """isEarningsDateEstimate"":true") > 0)
    If nd <> "" Then
        If IsDate(nd) Then YahooCalendar = True
    End If
    If lastCall <> "" Then
        If Not IsDate(lastCall) Then lastCall = ""
    End If
    Exit Function
Bad:
    YahooCalendar = False
End Function

Private Function YGet(ByVal url As String, ByVal cookie As String, ByRef status As Long, ByRef headers As String) As String
    Dim h As Object
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.Option(6) = False                                        ' WinHttpRequestOption_EnableRedirects
    h.SetTimeouts 5000, 5000, 10000, 10000
    h.Open "GET", url, False
    h.SetRequestHeader "User-Agent", UA
    If cookie <> "" Then h.SetRequestHeader "Cookie", cookie
    h.Send
    status = h.status
    headers = h.GetAllResponseHeaders
    YGet = h.ResponseText
End Function

Private Function CookiesFrom(ByVal headers As String) As String
    Dim ln As Variant, v As String, p As Long, out As String
    For Each ln In Split(headers, vbCrLf)
        If LCase$(Left$(ln, 11)) = "set-cookie:" Then
            v = Trim$(Mid$(ln, 12))
            p = InStr(v, ";")
            If p > 0 Then v = Left$(v, p - 1)
            If v <> "" Then
                If out <> "" Then out = out & "; "
                out = out & v
            End If
        End If
    Next ln
    CookiesFrom = out
End Function

Private Function UrlEnc(ByVal s As String) As String
    s = Replace(s, "%", "%25"): s = Replace(s, "/", "%2F")
    s = Replace(s, "+", "%2B"): s = Replace(s, "=", "%3D")
    UrlEnc = s
End Function

' First "fmt":"yyyy-mm-dd" inside the array that follows key.
Private Function FmtAfter(ByVal body As String, ByVal key As String) As String
    Dim p As Long: p = InStr(body, key)
    If p = 0 Then Exit Function
    Dim q As Long: q = InStr(p, body, """fmt"":""")
    Dim e As Long: e = InStr(p, body, "]")
    If q = 0 Then Exit Function
    If e > 0 And q > e Then Exit Function
    FmtAfter = Mid$(body, q + 7, 10)
End Function

' ----------------------------------------------------------------
'  Quarterly snapshot table
' ----------------------------------------------------------------
Private Function RenderTable(ws As Worksheet, ByVal mkt As String, prices As Object, ByVal srcNote As String) As Long
    Dim n As Long: n = m_n
    Dim i As Long
    Dim rev As Variant, eps As Variant, gp As Variant, opInc As Variant, ni As Variant, cogs As Variant, da As Variant
    Dim capex As Variant, cfo As Variant, dps As Variant, intExp As Variant
    Dim shares As Variant, inv As Variant, ar As Variant, ap As Variant, ca As Variant, cl As Variant
    Dim ltd As Variant, std As Variant, eq As Variant, cash As Variant, assets As Variant, liab As Variant, tax As Variant
    rev = m_v("Revenue"): eps = m_v("Eps"): gp = m_v("GrossProfit"): opInc = m_v("OperatingIncome")
    ni = m_v("NetIncome"): cogs = m_v("Cogs"): da = m_v("Da"): capex = m_v("CapEx"): cfo = m_v("Cfo")
    dps = m_v("Dividends"): intExp = m_v("InterestExpense")
    shares = m_v("Shares"): inv = m_v("Inventory"): ar = m_v("Ar"): ap = m_v("AccountsPayable")
    ca = m_v("CurrentAssets"): cl = m_v("CurrentLiabilities"): ltd = m_v("LongTermDebt"): std = m_v("ShortTermDebt")
    eq = m_v("StockholdersEquity"): cash = m_v("Cash"): assets = m_v("Assets"): liab = m_v("Liabilities")
    tax = m_v("EffectiveTaxRate")

    Dim px() As Variant, mcap() As Variant, sps() As Variant, pe() As Variant, ps() As Variant
    Dim revG() As Variant, epsG() As Variant, gm() As Variant, om() As Variant
    Dim ebitda() As Variant, ev() As Variant, evEb() As Variant, evS() As Variant, peg() As Variant
    Dim roe() As Variant, nm() As Variant, at() As Variant, em() As Variant, roic() As Variant
    Dim cr() As Variant, dio() As Variant, dso() As Variant, dpo() As Variant, ccc() As Variant
    Dim bvps() As Variant, cover() As Variant, fcf() As Variant, fcfC() As Variant, fcfY() As Variant
    Dim dy() As Variant, payout() As Variant
    ReDim px(1 To n): ReDim mcap(1 To n): ReDim sps(1 To n): ReDim pe(1 To n): ReDim ps(1 To n)
    ReDim revG(1 To n): ReDim epsG(1 To n): ReDim gm(1 To n): ReDim om(1 To n)
    ReDim ebitda(1 To n): ReDim ev(1 To n): ReDim evEb(1 To n): ReDim evS(1 To n): ReDim peg(1 To n)
    ReDim roe(1 To n): ReDim nm(1 To n): ReDim at(1 To n): ReDim em(1 To n): ReDim roic(1 To n)
    ReDim cr(1 To n): ReDim dio(1 To n): ReDim dso(1 To n): ReDim dpo(1 To n): ReDim ccc(1 To n)
    ReDim bvps(1 To n): ReDim cover(1 To n): ReDim fcf(1 To n): ReDim fcfC(1 To n): ReDim fcfY(1 To n)
    ReDim dy(1 To n): ReDim payout(1 To n)

    For i = 1 To n
        px(i) = Num(NearestPriceOnOrBefore(prices, m_end(i)))
        mcap(i) = Mu(px(i), shares(i))
        sps(i) = Dv(rev(i), shares(i))
        pe(i) = Dv(px(i), eps(i))
        ps(i) = Dv(mcap(i), rev(i))
        If i > 1 Then
            revG(i) = Growth(rev(i), rev(i - 1))
            epsG(i) = Growth(eps(i), eps(i - 1))
        Else
            revG(i) = "": epsG(i) = ""
        End If
        gm(i) = Dv(gp(i), rev(i))
        om(i) = Dv(opInc(i), rev(i))
        ebitda(i) = Ad(opInc(i), da(i))
        Dim debt As Variant: debt = ""
        If HasVal(ltd(i)) Then
            debt = CDbl(ltd(i))
            If HasVal(std(i)) Then debt = debt + CDbl(std(i))
        End If
        ev(i) = Sb(Ad(mcap(i), debt), cash(i))
        evEb(i) = Dv(ev(i), ebitda(i))
        evS(i) = Dv(ev(i), rev(i))
        peg(i) = ""
        If HasVal(pe(i)) And Pos(epsG(i)) Then peg(i) = CDbl(pe(i)) / (CDbl(epsG(i)) * 100)
        roe(i) = Dv(ni(i), eq(i))
        nm(i) = Dv(ni(i), rev(i))
        at(i) = Dv(rev(i), assets(i))
        em(i) = Dv(assets(i), eq(i))
        Dim t As Double: t = 0
        If HasVal(tax(i)) Then t = CDbl(tax(i))
        roic(i) = Dv(Mu(opInc(i), 1 - t), Ad(eq(i), ltd(i)))
        cr(i) = Dv(ca(i), cl(i))
        dio(i) = Mu(Dv(inv(i), cogs(i)), 91)
        dso(i) = Mu(Dv(ar(i), rev(i)), 91)
        dpo(i) = Mu(Dv(ap(i), cogs(i)), 91)
        ccc(i) = Sb(Ad(dio(i), dso(i)), dpo(i))
        bvps(i) = Dv(eq(i), shares(i))
        cover(i) = Dv(opInc(i), intExp(i))
        fcf(i) = Sb(cfo(i), capex(i))
        fcfC(i) = Dv(fcf(i), ni(i))
        fcfY(i) = Dv(fcf(i), mcap(i))
        dy(i) = Dv(dps(i), px(i))
        payout(i) = ""
        If HasVal(dps(i)) And Pos(eps(i)) Then payout(i) = CDbl(dps(i)) / CDbl(eps(i))
    Next i

    ' ---- write ----
    Dim unit As String: unit = IIf(mkt = "TW", "NTD", "USD")
    Call Section(ws, PG_TBL_TITLE, "QUARTERLY SNAPSHOT  .  last " & n & " quarters, oldest -> newest  .  (M) = " & unit & " millions")
    Dim r As Long: r = PG_TBL_HDR
    ws.cells(r, 1).Value = "METRIC"
    For i = 1 To n
        With ws.cells(r, i + 1)
            .Value = m_lbl(i)
            If m_q4(i) Then
                .Font.Italic = True
                .AddComment "Derived: full-year report minus Q1-Q3 (no standalone Q4 figure is filed). Balance-sheet rows are the year-end values; tax rate blank."
            End If
        End With
    Next i
    With ws.Range(ws.cells(r, 1), ws.cells(r, n + 1))
        .Font.Color = RR4_ACCENT: .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RR4_LINE
    End With
    r = r + 1
    Call TextRow(ws, r, "Period end", m_end)
    Call TextRow(ws, r, "Filed" & IIf(mkt = "TW", " (deadline)", ""), m_filed)
    Call TextRow(ws, r, "Form", m_form)

    Dim banners As New Collection
    Call Banner(ws, r, "MARKET DATA", banners)
    Call WRow(ws, r, "Stock Price", px, "price")
    Call WRow(ws, r, "Market Cap (M)", mcap, "money")
    Call WRow(ws, r, "Common Shares (M)", shares, "money")

    Call Banner(ws, r, "INCOME STATEMENT", banners)
    Call WRow(ws, r, "Revenue (M)", rev, "money")
    Call WRow(ws, r, "Sales Growth % (QoQ)", revG, "pct")
    Call WRow(ws, r, "Gross Profit (M)", gp, "money")
    Call WRow(ws, r, "Gross Margin", gm, "pct")
    Call WRow(ws, r, "EPS (GAAP)", eps, "ratio")
    Call WRow(ws, r, "EPS Growth % (QoQ)", epsG, "pct")
    Call WRow(ws, r, "Operating Income (M)", opInc, "money")
    Call WRow(ws, r, "Operating Margin", om, "pct")
    Call WRow(ws, r, "Net Income (M)", ni, "money")
    Call WRow(ws, r, "EBITDA (M)", ebitda, "money")

    Call Banner(ws, r, "VALUATION MULTIPLES  (on single-quarter figures)", banners)
    Call WRow(ws, r, "P/E", pe, "ratio")
    Call WRow(ws, r, "P/S", ps, "ratio")
    Call WRow(ws, r, "Sales per Share", sps, "price")
    Call WRow(ws, r, "EV/EBITDA", evEb, "ratio")
    Call WRow(ws, r, "EV/Sales", evS, "ratio")
    Call WRow(ws, r, "PEG Ratio", peg, "ratio")

    Call Banner(ws, r, "PROFITABILITY & RETURNS", banners)
    Call WRow(ws, r, "ROE", roe, "pct")
    Call WRow(ws, r, "  Net Margin (DuPont)", nm, "pct")
    Call WRow(ws, r, "  Asset Turnover (DuPont)", at, "ratio")
    Call WRow(ws, r, "  Equity Multiplier (DuPont)", em, "ratio")
    Call WRow(ws, r, "ROIC", roic, "pct")
    Call WRow(ws, r, "Effective Tax Rate", tax, "pct")

    Call Banner(ws, r, "WORKING CAPITAL", banners)
    Call WRow(ws, r, "Inventory (M)", inv, "money")
    Call WRow(ws, r, "Accounts Receivable (M)", ar, "money")
    Call WRow(ws, r, "Accounts Payable (M)", ap, "money")
    Call WRow(ws, r, "Current Assets (M)", ca, "money")
    Call WRow(ws, r, "Current Liabilities (M)", cl, "money")
    Call WRow(ws, r, "Current Ratio", cr, "ratio")
    Call WRow(ws, r, "Days Inventory Outstanding", dio, "days")
    Call WRow(ws, r, "Days Sales Outstanding", dso, "days")
    Call WRow(ws, r, "Days Payable Outstanding", dpo, "days")
    Call WRow(ws, r, "Cash Conversion Cycle (days)", ccc, "days")

    Call Banner(ws, r, "CAPITAL STRUCTURE", banners)
    Call WRow(ws, r, "Cash (M)", cash, "money")
    Call WRow(ws, r, "LT Debt (M)", ltd, "money")
    Call WRow(ws, r, "Short-Term Debt (M)", std, "money")
    Call WRow(ws, r, "Total Assets (M)", assets, "money")
    Call WRow(ws, r, "Total Liabilities (M)", liab, "money")
    Call WRow(ws, r, "Stockholders Equity (M)", eq, "money")
    Call WRow(ws, r, "Book Value/Share", bvps, "price")
    Call WRow(ws, r, "Enterprise Value (M)", ev, "money")
    Call WRow(ws, r, "Interest Expense (M)", intExp, "money")
    Call WRow(ws, r, "Interest Coverage (x)", cover, "ratio")

    Call Banner(ws, r, "CASH FLOW", banners)
    Call WRow(ws, r, "CFO (M)", cfo, "money")
    Call WRow(ws, r, "CapEx (M)", capex, "money")
    Call WRow(ws, r, "D&A (M)", da, "money")
    Call WRow(ws, r, "Free Cash Flow (M)", fcf, "money")
    Call WRow(ws, r, "FCF Conversion (FCF/NI)", fcfC, "ratio")
    Call WRow(ws, r, "FCF Yield", fcfY, "pct")

    Call Banner(ws, r, "DIVIDENDS", banners)
    Call WRow(ws, r, "Dividend per Share", dps, "price")
    Call WRow(ws, r, "Dividend Yield", dy, "pct")
    Call WRow(ws, r, "Payout Ratio", payout, "pct")

    Dim lastRow As Long: lastRow = r - 1
    ws.Range(ws.cells(PG_TBL_HDR + 1, 2), ws.cells(lastRow, n + 1)).HorizontalAlignment = xlRight
    ws.Range(ws.cells(PG_TBL_HDR, 2), ws.cells(PG_TBL_HDR, n + 1)).HorizontalAlignment = xlRight
    For i = 1 To n
        If m_q4(i) Then ws.Range(ws.cells(PG_TBL_HDR + 1, i + 1), ws.cells(lastRow, i + 1)).Font.Italic = True
    Next i
    Dim b As Variant
    For Each b In banners
        With ws.Range(ws.cells(CLng(b), 1), ws.cells(CLng(b), n + 1))
            .Interior.Color = CLR_BANNER
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .Font.Italic = False
        End With
    Next b

    ' ---- P/E and P/S range band ----
    Dim bc As Long: bc = n + 3
    ws.cells(PG_TBL_HDR, bc).Value = "RANGE (" & n & "Q)"
    ws.cells(PG_TBL_HDR, bc + 1).Value = "MIN"
    ws.cells(PG_TBL_HDR, bc + 2).Value = "MEDIAN"
    ws.cells(PG_TBL_HDR, bc + 3).Value = "MAX"
    With ws.Range(ws.cells(PG_TBL_HDR, bc), ws.cells(PG_TBL_HDR, bc + 3))
        .Font.Color = RR4_ACCENT: .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RR4_LINE
    End With
    Call BandRow(ws, PG_TBL_HDR + 1, bc, "P/E", pe)
    Call BandRow(ws, PG_TBL_HDR + 2, bc, "P/S", ps)

    Call WriteFlag(ws, lastRow + 1, srcNote)
    RenderTable = lastRow + 1
End Function

Private Sub Section(ws As Worksheet, ByVal r As Long, ByVal txt As String)
    With ws.cells(r, 1)
        .Value = txt
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 10
    End With
    ws.Rows(r).RowHeight = 20
End Sub

Private Sub Banner(ws As Worksheet, ByRef r As Long, ByVal txt As String, banners As Collection)
    ws.cells(r, 1).Value = txt
    banners.Add r
    r = r + 1
End Sub

Private Sub TextRow(ws As Worksheet, ByRef r As Long, ByVal label As String, vals() As String)
    ws.cells(r, 1).Value = label
    ws.cells(r, 1).Font.Color = CLR_MUTED
    Dim i As Long
    For i = 1 To m_n
        With ws.cells(r, i + 1)
            .NumberFormat = "@"
            .Value = vals(i)
            .Font.Color = CLR_MUTED
        End With
    Next i
    r = r + 1
End Sub

Private Sub WRow(ws As Worksheet, ByRef r As Long, ByVal label As String, vals As Variant, ByVal fmt As String)
    ws.cells(r, 1).Value = label
    Dim i As Long
    For i = 1 To m_n
        If HasVal(vals(i)) Then
            ws.cells(r, i + 1).Value = CDbl(vals(i))
        Else
            ws.cells(r, i + 1).Value = "-"
            ws.cells(r, i + 1).Font.Color = CLR_MUTED
        End If
    Next i
    Dim rng As Range: Set rng = ws.Range(ws.cells(r, 2), ws.cells(r, m_n + 1))
    Select Case fmt
        Case "money": rng.NumberFormat = "#,##0,,""M"";-#,##0,,""M"""
        Case "price": rng.NumberFormat = "#,##0.00"
        Case "ratio": rng.NumberFormat = "0.00"
        Case "days":  rng.NumberFormat = "0.0"
        Case "pct":   rng.NumberFormat = "0.00%"
    End Select
    r = r + 1
End Sub

Private Sub BandRow(ws As Worksheet, ByVal r As Long, ByVal bc As Long, ByVal label As String, vals As Variant)
    ws.cells(r, bc).Value = label
    Dim cnt As Long, i As Long, a() As Double
    ReDim a(1 To m_n)
    For i = 1 To m_n
        If HasVal(vals(i)) Then cnt = cnt + 1: a(cnt) = CDbl(vals(i))
    Next i
    If cnt = 0 Then Exit Sub
    ReDim Preserve a(1 To cnt)
    ws.cells(r, bc + 1).Value = Application.WorksheetFunction.Min(a)
    ws.cells(r, bc + 2).Value = Application.WorksheetFunction.Median(a)
    ws.cells(r, bc + 3).Value = Application.WorksheetFunction.Max(a)
    ws.Range(ws.cells(r, bc + 1), ws.cells(r, bc + 3)).NumberFormat = "0.00"
End Sub

Private Sub WriteFlag(ws As Worksheet, ByVal r As Long, ByVal txt As String)
    With ws.cells(r, 1)
        .Value = txt
        .Font.Color = CLR_FLAG
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

' ----------------------------------------------------------------
'  TW monthly revenue
' ----------------------------------------------------------------
Private Function RenderMonthly(ws As Worksheet, ByVal atRow As Long, ByVal coId As String) As Long
    Dim mr As Object
    Set mr = modMOPSData.GetTwMonthlyRevenue(coId, TW_MONTHS)
    If Not mr("ok") Then
        Call WriteFlag(ws, atRow, "Monthly revenue unavailable: " & CStr(mr("error")))
        RenderMonthly = atRow
        Exit Function
    End If
    Dim labels As Variant, rev As Variant, yoy As Variant
    labels = mr("labels"): rev = mr("rev"): yoy = mr("yoy")
    Dim n As Long: n = UBound(labels) - LBound(labels) + 1
    Call Section(ws, atRow, "MONTHLY REVENUE  .  last " & n & " months, oldest -> newest  .  NTD millions")
    Dim h As Long: h = atRow + 1
    ws.cells(h, 1).Value = "MONTH"
    ws.cells(h + 1, 1).Value = "Revenue (M)"
    ws.cells(h + 2, 1).Value = "YoY %"
    Dim j As Long, c As Long
    For j = LBound(labels) To UBound(labels)
        c = 2 + (j - LBound(labels))
        ws.cells(h, c).NumberFormat = "@"
        ws.cells(h, c).Value = CStr(labels(j))
        If HasVal(rev(j)) Then
            ws.cells(h + 1, c).Value = CDbl(rev(j)) / 1000#          ' t21sc03 is in NTD thousands
            ws.cells(h + 1, c).NumberFormat = "#,##0"
        Else
            ws.cells(h + 1, c).Value = "-"
        End If
        If HasVal(yoy(j)) Then
            ws.cells(h + 2, c).Value = CDbl(yoy(j))
            ws.cells(h + 2, c).NumberFormat = "0.0%"
            If CDbl(yoy(j)) >= 0 Then ws.cells(h + 2, c).Font.Color = RGB(220, 80, 80) Else ws.cells(h + 2, c).Font.Color = RGB(80, 200, 120)
        Else
            ws.cells(h + 2, c).Value = "-"
        End If
    Next j
    With ws.Range(ws.cells(h, 1), ws.cells(h, n + 1))
        .Font.Color = RR4_ACCENT: .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RR4_LINE
    End With
    ws.Range(ws.cells(h, 2), ws.cells(h + 2, n + 1)).HorizontalAlignment = xlRight
    Call WriteFlag(ws, h + 3, "Source: MOPS t21sc03 (single month recomputed from cumulative). The newest months can post-date the latest quarterly report.")
    RenderMonthly = h + 3
End Function

' ----------------------------------------------------------------
' Sheet event code (a copy of RR4/SheetEarnings_Code.txt), written into the
' document module the first time the page is built.  Needs "Trust access to
' the VBA project object model", like the RRG / Thesis pages.
Private Sub EnsureSheetCode(ws As Worksheet)
    On Error GoTo Skip
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 Then
            If comp.Properties("Name").Value = ws.Name Then
                Dim cm As Object: Set cm = comp.CodeModule
                If cm.CountOfLines > 0 Then
                    If InStr(cm.Lines(1, cm.CountOfLines), "EarningsChange") > 0 Then Exit Sub
                    cm.DeleteLines 1, cm.CountOfLines
                End If
                cm.AddFromString "Option Explicit" & vbCrLf & vbCrLf & _
                    "' Earnings page: TICKER / MKT input -> modEarnings.ShowEarnings (see modEarnings.bas)" & vbCrLf & _
                    "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
                    "    Call EarningsChange(Me, Target)" & vbCrLf & _
                    "End Sub" & vbCrLf
                Exit Sub
            End If
        End If
    Next comp
    Exit Sub
Skip:
    Call NavNotify("Earnings page built, but its sheet event code could not be written (enable Trust access to the VBA project object model, or paste RR4/SheetEarnings_Code.txt)", True)
End Sub
