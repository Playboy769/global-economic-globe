Attribute VB_Name = "PortfolioDashboard_v3"
Option Explicit

' ================================================================
'  PORTFOLIO DASHBOARD v3.0 - Bloomberg Terminal Style
' ================================================================

Private Const SH_PORT  As String = "RR4"
Private Const SH_TRANS As String = "Transactions"
Private Const SH_REAL  As String = "Realized"
Private Const SH_HIST  As String = "HistoryLog"

' Raw-data sheet holding the YTD baselines (see HISTORYLOG LAYOUT v2 below).
' Kept off HistoryLog on purpose: the first version parked the baselines in
' P1:R6, right next to the leftovers of the old layout, and the two blocks
' were impossible to tell apart on screen.
Private Const SH_HRAW     As String = "HistoryRaw"
Private Const HL_BASE_ROW As Long = 3   ' first baseline row on HistoryRaw

' ================================================================
'  RR4 PAGE LAYOUT v4.1 (2026-09-12; v4 2026-09-11)
' ----------------------------------------------------------------
'  Rows 1-3                    nav bar (modNav.DrawNavRows, code "P")
'  Upper-left  A:G  rows 4-22  total, USD/TWD, ARRANGE <GO>, weight bar,
'                              daily log, summary
'  Upper-right I:R  rows 4-24  ticker panel (TickerInsight module)
'  (row 25 blank)
'  Chart band       rows 26-39 weight donut at F (RR4_DONUT) and the
'                              realized-PnL line across J:S (RR4_RLPNL)
'                              (v4.7 - both used to sit right of the panel)
'  (row 40 blank)
'  Position log                title row 41, headers row 42, data from 43
'  (row numbers in the notes below are the v4 ones; add RR4_TOP = 3)
'
'  Position log columns (BROKER column and broker group rows removed):
'    A TICKER   B NAME      C ENTRY DT  D DAYS      E SECTOR   F NET EXPOS
'    G SHARES   H ENTRY PX  I LAST      J % CHG     K UNRL PNL L WT%
'    M W.BETA   N BETA 180D O P.TARGET  P SWING RISK (typed by hand)
'    T (hidden) default-order key, used by ARRANGE "DEF"
'
'  ARRANGE <GO> (D2): UNU/UND = UNRL PNL, PCU/PCD = % CHG, DAU/DAD = DAYS,
'  WTU/WTD = WT%.  U = low -> high, D = high -> low.  Blank or DEF = TW
'  codes first (numeric), then US tickers A-Z.  Typing a code re-sorts the
'  rows already on the sheet (no price refetch) - see ApplyArrange.
'
'  ResetSheetStyle clears the whole sheet, so every hand-entered value is
'  read FIRST and written back: B2 USD/TWD, D2 ARRANGE code, S1/S2 (the
'  InceptionDate / StartingCapital names set by SetupPortfolioConfig), the
'  SWING RISK column, and the ticker panel's J1 ticker / K17 price target.
'  Up to v3, B2 and S1:S2 were read AFTER the clear, so the rate always
'  fell back to 31.6 and inception / capital to their hard-coded defaults.
' ================================================================
' v4.1 (2026-09-12): rows 1-3 now hold the nav bar (modNav) and the whole
' page above moved down by RR4_TOP - every row number in this module that
' belongs to the page is written as RR4_TOP + <its v4 row>, and the cell
' constants below are the moved addresses. Other modules must read the
' page through these constants / RR4FxRate(), never by literal address.
Public Const RR4_TOP       As Long = 3
' v4.2 (2026-09-12): column A is a blank spacer for breathing room, so the
' page also moved one column right - body columns are B:Q, the ticker panel
' J:S, the hidden order key V, and the config cells T1/T2 (they were S1/S2,
' which the panel now covers; RebuildPortfolioDashboard migrates them).
Public Const RR4_LEFT      As Long = 1
Public Const RR4_CFG_INC   As String = "T1"
Public Const RR4_CFG_CAP   As String = "T2"
Public Const RR4_TOTAL_CELL As String = "B4"
Public Const RR4_FX_CELL   As String = "C5"
Public Const RR4_ARR_CELL  As String = "E5"
' v4.7 (2026-09-12): a 14-row chart band (RR4_CHART_TOP..) sits between the
' upper blocks and the position log, which moved down from 27/28/29.
' v4.7.1: one blank row above (25) and below (40) the band; the ticker
' panel was shortened to row 24 to make room (TickerInsight TI_BOTTOM).
Public Const RR4_CHART_TOP As Long = 26
Public Const RR4_CHART_ROWS As Long = 14
' WATCHLIST (v4.9, 2026-09-12; v4.9.1 entry-row flow): B:E of the chart
' band, left of the donut. Title row 25, header 26.
'   Row 27 = ENTRY ROW (input cells): type B ticker / C strategy / D entry
'            target; as soon as ticker AND target are both in, the sheet
'            code hands the row to WatchlistCommitEntry, which appends it
'            to the list and clears row 27 for the next one.
'   Rows 28-38 = the saved list (11 entries), E = live last price, a row
'            whose last <= target is lit. Double-click a saved row to
'            delete it (WatchlistDeleteRow, from Worksheet_BeforeDoubleClick).
' Saved rows survive the clear (ReadWatchlist / DrawWatchlist).
Public Const RR4_WL_TITLE  As Long = 25
Public Const RR4_WL_HDR    As Long = 26
Public Const RR4_WL_ENTRY  As Long = 27
Public Const RR4_WL_FIRST  As Long = 28
Public Const RR4_WL_LAST   As Long = 38
Public Const RR4_POS_TITLE As Long = 41
Public Const RR4_POS_HDR   As Long = 42
Public Const RR4_POS_FIRST As Long = 43
Private Const RR4_DONUT_NAME As String = "RR4_DONUT"
Private Const RR4_RLPNL_NAME As String = "RR4_RLPNL"   ' realized-PnL line chart (v4.6)
Private Const RR4_NCOL      As Long = 17    ' last body column, B:Q
Private Const RR4_ORD_COL   As Long = 22    ' V (hidden)
Private Const RR4_SWING_COL As Long = 17    ' Q
Private Const RR4_LOG_ROWS  As Long = 5     ' daily-log trade lines, rows 10-14
Private Const RR4_POS_ROW_H As Double = 24  ' position-log data rows (v4.5, was 18)
Private Const RR4_WBAR_PREFIX As String = "RR4W_"
' Every cell the user types into is painted this dark grey (RGB 40,40,40)
' with WHITE text (RR4_INPUT_FG): the nav command cell C1, USD/TWD C5,
' ARRANGE E5, SWING RISK (Q29:Q..), the ticker panel's K4 / L20, and the
' B2 input of the VT / CC pages. (v4.3, 2026-09-12: was 70,70,70 + yellow.)
Public Const RR4_INPUT_BG  As Long = 2631720
Public Const RR4_INPUT_FG  As Long = 16777215
' SWING RISK column only: a shade darker than the other input cells so it
' sits closer to the row stripes (v4.5.1, RGB 25,25,25).
Private Const RR4_SWING_BG As Long = 1644825
' Accent colour of the RR4 page and the nav bar (v4.5, 2026-09-12): dark
' orange RGB(200,100,0) - was the amber RGB(255,192,0) the other report
' pages (Vol / Corr) still use.
Public Const RR4_ACCENT    As Long = 25800
' Section divider lines on the RR4 page (nav bar bottom, USD/TWD row, TODAY
' row, position-log title / header / last row, ticker-panel history header)
' are this dark grey since v4.3 - was the amber RGB(255,192,0).
Public Const RR4_LINE      As Long = 4605510    ' RGB(70,70,70)
' Donut slice palette (v4.3): one colour per position, cycled. Chosen to
' stay apart from each other on the black chart background.
Private Const RR4_PALETTE_N As Long = 12

' ================================================================
'  MAIN ENTRY
' ================================================================
Sub RebuildPortfolioDashboard()
    Dim wsP  As Worksheet
    Dim wsTr As Worksheet
    Dim wsR  As Worksheet
    Set wsP = ThisWorkbook.Sheets(SH_PORT)
    Set wsTr = ThisWorkbook.Sheets(SH_TRANS)
    Set wsR = ThisWorkbook.Sheets(SH_REAL)

    If Not g_PriceCache Is Nothing Then
        g_PriceCache.RemoveAll
        g_CacheTime = Now
    End If

    ' The page writes into its own input cells (B1, D5, J4, K20) - keep the
    ' Worksheet_Change / Workbook_SheetChange handlers out of the way.
    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fail

    ' --- hand-entered values: read BEFORE the sheet is cleared ---
    Dim exRate As Double: exRate = GetExRate(wsP)
    Dim arrCode As String: arrCode = UCase(CellStr(wsP.Range(RR4_ARR_CELL).Value))
    Dim tiTicker As String: tiTicker = UCase(CellStr(wsP.Range(TI_TICKER_CELL).Value))
    Dim tiTarget As Variant: tiTarget = wsP.Range(TI_TARGET_CELL).Value
    ' config (SetupPortfolioConfig): T1/T2 since v4.2, S1/S2 before it - the
    ' panel covers S now, so carry the old pair over once
    Dim cfgInc As Variant: cfgInc = wsP.Range(RR4_CFG_INC).Value
    Dim cfgCap As Variant: cfgCap = wsP.Range(RR4_CFG_CAP).Value
    If Not IsDate(cfgInc) Then cfgInc = wsP.Range("S1").Value
    If NumOr0(cfgCap) <= 0 Then cfgCap = wsP.Range("S2").Value
    Dim swingRiskMap As Object: Set swingRiskMap = ReadSwingRisk(wsP)
    Dim wl As Variant: wl = ReadWatchlist(wsP)
    ' A sheet still on an older layout has other data sitting in these cells:
    ' keep only a real ARRANGE code and a numeric target for a real ticker.
    If Not IsArrangeCode(arrCode) Then arrCode = ""
    ' only trust the panel cells when the panel label is actually there -
    ' on an older layout those cells hold a column header, not a ticker
    If CellStr(wsP.cells(RR4_TOP + 1, 10).Value) <> "TICKER <GO>" Then tiTicker = ""
    If IsError(tiTarget) Then tiTarget = Empty
    If tiTicker = "" Or Not IsNumeric(tiTarget) Then tiTarget = Empty

    Call ResetSheetStyle(wsP)
    If IsDate(cfgInc) Then wsP.Range(RR4_CFG_INC).Value = cfgInc
    If NumOr0(cfgCap) > 0 Then wsP.Range(RR4_CFG_CAP).Value = cfgCap
    Call PointConfigNames(wsP)
    Call RemoveLegacyButtons(wsP)
    Call DrawNavRows(wsP, "P")

    Dim positions As Object
    Set positions = BuildPositions(wsTr)

    Dim posData()    As Variant
    Dim totalMktTWD  As Double
    Dim totalCostTWD As Double
    Dim totalUnrlTWD As Double
    Dim posCount     As Long

    Call CalcPositions(positions, exRate, posData, totalMktTWD, totalCostTWD, totalUnrlTWD, posCount)
    Call CalculateRealizedPnL

    Dim realPnL As Double
    On Error Resume Next
    ' PNL(TWD) moved G -> H when RET% was inserted as column B (Attach.bas)
    realPnL = Application.WorksheetFunction.Sum(wsR.Columns("H"))
    On Error GoTo Fail

    Dim portBeta As Double
    portBeta = CalcPortBeta(posData)

    ' previous trading day's cumulative PnL - read before LogHistory
    ' rewrites / appends today's HistoryLog row
    Dim prevPnL As Variant: prevPnL = PrevDayCumPnL()

    Call DrawHeader(wsP, totalMktTWD, exRate, arrCode)
    Call DrawDailyLog(wsP, posData, exRate, totalMktTWD, totalUnrlTWD + realPnL, realPnL, prevPnL)
    Call DrawSummary(wsP, posData, totalMktTWD, totalCostTWD, totalUnrlTWD, realPnL, portBeta, posCount)
    Call DrawColumnHeaders(wsP)
    Dim lastDataRow As Long
    lastDataRow = WritePositionRows(wsP, posData, totalMktTWD, swingRiskMap)
    Call ApplyArrange(arrCode)
    Call DrawWatchlist(wsP, wl)
    Call DrawDisclaimer(wsP, lastDataRow)
    Call RenderTickerPanel(tiTicker, tiTarget)
    Call LogHistory(totalMktTWD, totalUnrlTWD + realPnL, realPnL)
    Call DrawRealizedChart(wsP)     ' after LogHistory, so today's row is on the line

    Application.ScreenUpdating = True
    ' (v4.6: the separate "Analysis" sheet / DrawDeepAnalysis is gone - its
    ' currency split lives in the Summary, the rest was not used)
    Application.EnableEvents = prevEvents
    Application.StatusBar = "Dashboard updated: " & Format(Now, "hh:mm:ss")
    Call NavNotify("UP done " & Format(Now, "hh:mm:ss") & " - " & posCount & " positions")
    Exit Sub

Fail:
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = True
    Application.StatusBar = "Dashboard update FAILED: " & errD
    If Application.Visible Then MsgBox "RebuildPortfolioDashboard error " & errN & ": " & errD, vbCritical
End Sub

' ================================================================
'  HISTORYLOG LAYOUT v2 (2026-08-13) - every index return is YTD
' ----------------------------------------------------------------
'    A  Date                    B  TotalMarketValue
'    C  TotalCumulativePnL      D  Realized PnL
'    E  Realized PnL Daily Chg%
'    F  Price (SPY)             G  YTD Ret% (SPY)
'    H  Price (QQQ)             I  YTD Ret% (QQQ)
'    J  Price (TWII)            K  YTD Ret% (TWII)
'    L  Price (SOX)             M  YTD Ret% (SOX)
'    N onwards  unused - keep empty
'    Z3         GitHub token (Attach.bas) - never write to column Z
'
'  RAW DATA lives on its own sheet, "HistoryRaw":
'    A  Ticker | B  Baseline date | C  Baseline close | D  Fetched at
'    rows 3-6 = SPY / QQQ / ^TWII / ^SOX, rebuilt when the year rolls over.
'  Every Ret% cell divides by HistoryRaw!$C$n, so each index is measured
'  against its OWN first-trading-day close and the raw inputs stay visible.
'
'  Ret% = (price - close of the FIRST TRADING DAY of the current year)
'         / that close.  It is NOT relative to row 2 any more.
'  ^TWOII (OTC index) was dropped - it never returned data from Yahoo.
'  A price that fails to download is left BLANK, never written as 0
'  (a 0 used to turn into a -100% return and wreck the charts).
'
'  To rebuild existing rows (wrong columns / stale prices): RepairHistoryLog.
' ================================================================
Private Function HistTickers() As Variant
    HistTickers = Array("SPY", "QQQ", "^TWII", "^SOX")
End Function

Private Function HistPriceCols() As Variant
    HistPriceCols = Array("F", "H", "J", "L")
End Function

Private Function HistRetCols() As Variant
    HistRetCols = Array("G", "I", "K", "M")
End Function

Private Sub LogHistory(totalMkt As Double, totalPnL As Double, realPnL As Double)
    Dim wsH As Worksheet
    On Error Resume Next
    Set wsH = ThisWorkbook.Sheets(SH_HIST)
    On Error GoTo 0
    If wsH Is Nothing Then Exit Sub

    Call EnsureHistoryHeaders(wsH)
    Call EnsureYTDBaselines(False)

    Dim nr As Long: nr = wsH.cells(wsH.Rows.count, "A").End(xlUp).row + 1
    If nr > 2 Then
        ' guard the date: text or an error value in column A used to abort the
        ' whole rebuild here, after the page had already been drawn
        Dim lastStamp As Variant: lastStamp = wsH.cells(nr - 1, 1).Value
        If IsDate(lastStamp) Then
            If Int(CDate(lastStamp)) = Date Then nr = nr - 1
        End If
    End If

    With wsH
        .cells(nr, "A").Value = Now
        .cells(nr, "B").Value = totalMkt
        .cells(nr, "C").Value = totalPnL
        .cells(nr, "D").Value = realPnL

        .cells(nr, "A").NumberFormat = "yyyy/m/d h:mm:ss"
        .Range("B" & nr & ":D" & nr).NumberFormat = "#,##0"
    End With

    Dim tickers As Variant, priceCols As Variant
    tickers = HistTickers()
    priceCols = HistPriceCols()

    Dim i As Long, px As Double
    For i = LBound(tickers) To UBound(tickers)
        px = 0
        On Error Resume Next
        px = GetStockPrice(CStr(tickers(i)))
        On Error GoTo 0
        If px > 0 Then
            wsH.cells(nr, priceCols(i)).Value = px
        Else
            ' download failed - leave the cell blank instead of logging a 0
            wsH.cells(nr, priceCols(i)).ClearContents
        End If
        wsH.cells(nr, priceCols(i)).NumberFormat = "#,##0.00"
    Next i

    Call WriteHistoryRowFormulas(wsH, nr)
    Call RebuildRealizedHistory(wsH)
End Sub

' ================================================================
'  Column D of every HistoryLog row = realized PnL accumulated up to that
'  row's date, taken from the Realized sheet AS IT IS NOW (v4.8, 2026-09-12).
'  Up to v4.7 the column was an append-only snapshot of SUM(Realized!H) at
'  the moment of each UP, which drifted away from the sheet in two ways:
'  the USD legs were converted at that day's USD/TWD (the Realized sheet
'  is re-converted at the current rate on every UP), and a trade entered
'  or re-dated after the fact re-pairs the FIFO lots for earlier days
'  (2026-09-08/09 were off by -2,070 / +2,373 for exactly that reason,
'  and showed up as a fake spike on the REALIZED PNL line). Restating the
'  column from the Realized sheet keeps D, its Daily Chg% (E) and the
'  chart consistent with the realized table at all times. Column C
'  (total cumulative PnL) is left as a snapshot - its unrealized part
'  cannot be restated.
'  Realized sheet layout (Attach.CalculateRealizedPnL): H = PNL(TWD),
'  I = exit date.
' ================================================================
Public Sub RebuildRealizedHistory(Optional ByVal wsH As Worksheet = Nothing)
    If wsH Is Nothing Then
        On Error Resume Next: Set wsH = ThisWorkbook.Sheets(SH_HIST): On Error GoTo 0
        If wsH Is Nothing Then Exit Sub
    End If
    Dim wsR As Worksheet
    On Error Resume Next: Set wsR = ThisWorkbook.Sheets(SH_REAL): On Error GoTo 0
    If wsR Is Nothing Then Exit Sub

    ' realized trades: (exit date, PnL TWD)
    Dim lastR As Long: lastR = wsR.cells(wsR.Rows.count, "A").End(xlUp).row
    Dim n As Long, r As Long
    Dim exD() As Long, exP() As Double
    If lastR >= 2 Then
        ReDim exD(1 To lastR - 1): ReDim exP(1 To lastR - 1)
        For r = 2 To lastR
            Dim dv As Variant: dv = wsR.cells(r, "I").Value
            If IsDate(dv) Then
                n = n + 1
                exD(n) = Int(CDbl(CDate(dv)))
                exP(n) = NumOr0(wsR.cells(r, "H").Value)
            End If
        Next r
    End If

    Dim lastH As Long: lastH = wsH.cells(wsH.Rows.count, "A").End(xlUp).row
    For r = 2 To lastH
        Dim hv As Variant: hv = wsH.cells(r, "A").Value
        If IsDate(hv) Then
            Dim dayN As Long: dayN = Int(CDbl(CDate(hv)))
            Dim cum As Double: cum = 0
            Dim k As Long
            For k = 1 To n
                If exD(k) <= dayN Then cum = cum + exP(k)
            Next k
            If Abs(NumOr0(wsH.cells(r, "D").Value) - cum) > 0.005 Then wsH.cells(r, "D").Value = cum
        End If
    Next r
End Sub

' Header row - rewritten every run so the layout stays self-describing.
' Only A1:M1 plus the P1:R6 baseline block are touched; column Z is left alone.
Private Sub EnsureHistoryHeaders(wsH As Worksheet)
    Dim hdr As Variant
    hdr = Array("Date", "TotalMarketValue", "TotalCumulativePnL", _
                "Realized PnL", "Realized PnL Daily Chg%", _
                "Price (SPY)", "YTD Ret% (SPY)", _
                "Price (QQQ)", "YTD Ret% (QQQ)", _
                "Price (TWII)", "YTD Ret% (TWII)", _
                "Price (SOX)", "YTD Ret% (SOX)")
    Dim i As Long
    For i = LBound(hdr) To UBound(hdr)
        If CStr(wsH.cells(1, i + 1).Value) <> CStr(hdr(i)) Then
            wsH.cells(1, i + 1).Value = hdr(i)
        End If
    Next i
End Sub

' The HistoryRaw sheet, created on first use.
Private Function GetRawSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_HRAW)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = SH_HRAW
    End If
    Set GetRawSheet = ws
End Function

' YTD baselines on HistoryRaw - close of the first trading day of the year.
' Rebuilt when the year rolls over, when a baseline is missing, or on demand.
Private Sub EnsureYTDBaselines(forceRefresh As Boolean)
    Dim wsR As Worksheet: Set wsR = GetRawSheet()
    Dim yr As Long: yr = Year(Date)
    Dim tickers As Variant: tickers = HistTickers()

    Dim needBuild As Boolean: needBuild = forceRefresh
    If CStr(wsR.Range("B1").Value) <> CStr(yr) Then needBuild = True

    Dim i As Long
    If Not needBuild Then
        For i = LBound(tickers) To UBound(tickers)
            Dim bv As Variant
            bv = wsR.cells(HL_BASE_ROW + i, "C").Value
            ' IsNumeric(Empty) is True in VBA, so test emptiness separately
            If IsEmpty(bv) Then
                needBuild = True
            ElseIf Not IsNumeric(bv) Then
                needBuild = True
            End If
        Next i
    End If
    If Not needBuild Then Exit Sub

    wsR.Range("A1").Value = "YTD BASELINE YEAR"
    wsR.Range("B1").Value = yr
    wsR.Range("B1").NumberFormat = "0"
    wsR.Range("A2").Value = "Ticker"
    wsR.Range("B2").Value = "Baseline date"
    wsR.Range("C2").Value = "First trading day close"
    wsR.Range("D2").Value = "Fetched at"

    Dim px As Double, bd As Date
    For i = LBound(tickers) To UBound(tickers)
        Dim rr As Long: rr = HL_BASE_ROW + i
        wsR.cells(rr, "A").Value = tickers(i)
        px = GetFirstTradingDayClose(CStr(tickers(i)), yr, bd)
        If px > 0 Then
            If bd > 0 Then
                wsR.cells(rr, "B").Value = bd
                wsR.cells(rr, "B").NumberFormat = "yyyy/m/d"
            End If
            wsR.cells(rr, "C").Value = px
            wsR.cells(rr, "C").NumberFormat = "#,##0.00"
            wsR.cells(rr, "D").Value = Now
            wsR.cells(rr, "D").NumberFormat = "yyyy/m/d h:mm"
        End If
    Next i

    wsR.Range("A:D").EntireColumn.AutoFit
End Sub

' Ret% formulas + realized-PnL daily change for one row.
Private Sub WriteHistoryRowFormulas(wsH As Worksheet, r As Long)
    Dim priceCols As Variant, retCols As Variant
    priceCols = HistPriceCols()
    retCols = HistRetCols()

    Dim i As Long, pc As String, rc As String, bc As String
    For i = LBound(retCols) To UBound(retCols)
        pc = priceCols(i) & r
        rc = retCols(i)
        ' each index divides by its OWN baseline on the raw-data sheet
        bc = SH_HRAW & "!$C$" & (HL_BASE_ROW + i)
        wsH.cells(r, rc).Formula = "=IFERROR(IF(" & pc & "="""","""",(" & pc & "-" & bc & ")/" & bc & "),"""")"
        wsH.cells(r, rc).NumberFormat = "0.00%"
    Next i

    ' Realized PnL daily change - "-" instead of #DIV/0! while realized PnL is 0
    If r > 2 Then
        wsH.cells(r, "E").Formula = "=IFERROR((D" & r & "-D" & (r - 1) & ")/ABS(D" & (r - 1) & "),""-"")"
    Else
        wsH.cells(r, "E").Formula = "=""-"""
    End If
    wsH.cells(r, "E").NumberFormat = "0.00%"
End Sub

' Close of the first trading day of `yr` for one ticker.
' Yahoo is asked for Jan 1 - Feb 15 and the first non-null close is taken,
' so New Year holidays / market closures do not matter.
Public Function GetFirstTradingDayClose(Ticker As String, yr As Long, ByRef outDate As Date) As Double
    GetFirstTradingDayClose = 0
    outDate = 0

    On Error GoTo ErrHandler
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim p1 As Long, p2 As Long
    p1 = DateDiff("s", #1/1/1970#, DateSerial(yr, 1, 1)) - 8 * 3600
    p2 = DateDiff("s", #1/1/1970#, DateSerial(yr, 2, 15))

    Dim safeTicker As String
    safeTicker = Replace(Ticker, "^", "%5E")

    Dim url As String
    url = "https://query2.finance.yahoo.com/v8/finance/chart/" & safeTicker & _
          "?period1=" & p1 & "&period2=" & p2 & "&interval=1d"

    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send

    Dim resp As String: resp = http.responseText

    Dim tsArr() As String, clArr() As String
    tsArr = Split(ExtractJsonArray(resp, """timestamp"":["), ",")
    clArr = Split(ExtractJsonArray(resp, """close"":["), ",")

    Dim i As Long, v As String
    For i = LBound(clArr) To UBound(clArr)
        v = Trim(clArr(i))
        If Len(v) > 0 And InStr(v, "null") = 0 And IsNumeric(v) Then
            GetFirstTradingDayClose = CDbl(v)
            If i <= UBound(tsArr) Then
                If IsNumeric(Trim(tsArr(i))) Then
                    outDate = DateAdd("s", CLng(Trim(tsArr(i))) + 8 * 3600, #1/1/1970#)
                End If
            End If
            Exit Function
        End If
    Next i

ErrHandler:
End Function

Private Function ExtractJsonArray(resp As String, key As String) As String
    Dim pos As Long: pos = InStr(resp, key)
    If pos = 0 Then Exit Function
    pos = pos + Len(key)
    Dim endPos As Long: endPos = InStr(pos, resp, "]")
    If endPos = 0 Then Exit Function
    ExtractJsonArray = Mid(resp, pos, endPos - pos)
End Function
Function GetHistoricalPrice(Ticker As String, targetDate As Date) As Double
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Unix epoch is UTC; TW market is UTC+8 (8h ahead) - opening crosses a day boundary
    Dim p1 As Long, p2 As Long
    p1 = DateDiff("s", #1/1/1970#, targetDate) - 8 * 3600
    p2 = p1 + 86400
    
    ' ^ -- URL encode -- %5E
    Dim safeTicker As String
    safeTicker = Replace(Ticker, "^", "%5E")
    
    Dim url As String
    url = "https://query2.finance.yahoo.com/v8/finance/chart/" & safeTicker & _
          "?period1=" & p1 & "&period2=" & p2 & "&interval=1d"
    
    On Error GoTo ErrHandler
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    
    Dim resp As String: resp = http.responseText
    
    ' ------------------------------------------------------------
    Dim pos As Long
    pos = InStr(resp, """close"":[")
    If pos = 0 Then GoTo ErrHandler
    pos = pos + 9
    Dim endPos As Long: endPos = InStr(pos, resp, "]")
    Dim closeVal As String: closeVal = Trim(Mid(resp, pos, endPos - pos))
    If Len(closeVal) = 0 Or InStr(closeVal, "null") > 0 Then GoTo ErrHandler
    
    GetHistoricalPrice = CDbl(Split(closeVal, ",")(0))
    Exit Function
    
ErrHandler:
    GetHistoricalPrice = 0
End Function
' Fill in any missing index price on existing rows (all four indices) and
' rewrite every Ret% / daily-change formula against the YTD baseline.
' Cells that already hold a price are left untouched.
Sub BackfillHistory()
    Dim wsH As Worksheet
    On Error Resume Next
    Set wsH = ThisWorkbook.Sheets(SH_HIST)
    On Error GoTo 0
    If wsH Is Nothing Then MsgBox "History sheet not found": Exit Sub

    Dim lastRow As Long
    lastRow = wsH.cells(wsH.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then MsgBox "No data to backfill": Exit Sub

    Call EnsureHistoryHeaders(wsH)
    Call EnsureYTDBaselines(False)

    Application.ScreenUpdating = False

    Dim tickers As Variant, priceCols As Variant
    tickers = HistTickers()
    priceCols = HistPriceCols()

    Dim filled() As Long
    ReDim filled(LBound(tickers) To UBound(tickers))

    Dim i As Long, k As Long
    For i = 2 To lastRow
        If wsH.cells(i, "A").Value = "" Then GoTo NextRow

        Dim targetDate As Date
        targetDate = 0
        On Error Resume Next
        targetDate = CDate(Int(CDbl(wsH.cells(i, "A").Value)))
        On Error GoTo 0
        If targetDate = 0 Then GoTo NextRow

        For k = LBound(tickers) To UBound(tickers)
            If wsH.cells(i, priceCols(k)).Value = "" Then
                Dim price As Double
                price = GetHistoricalPrice(CStr(tickers(k)), targetDate)
                If price > 0 Then
                    wsH.cells(i, priceCols(k)).Value = price
                    filled(k) = filled(k) + 1
                End If
            End If
            wsH.cells(i, priceCols(k)).NumberFormat = "#,##0.00"
        Next k

        Call WriteHistoryRowFormulas(wsH, i)

        Application.StatusBar = "Backfilling row " & i & "/" & lastRow
NextRow:
    Next i

    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim msg As String
    msg = "Backfill complete!" & vbLf & _
          "Ret% formulas rebuilt on YTD basis (" & lastRow - 1 & " rows)" & vbLf
    For k = LBound(tickers) To UBound(tickers)
        msg = msg & tickers(k) & " (col " & priceCols(k) & "): filled " & filled(k) & " gaps" & vbLf
    Next k
    msg = msg & "Prices that could not be downloaded were left blank."
    MsgBox msg
End Sub

' ================================================================
'  REPAIR: rebuild every index price column from Yahoo history
' ----------------------------------------------------------------
'  The first attempt at a migration never ran.  LogHistory rewrites the
'  header row on every update, so the sheet ended up showing layout-v2
'  headers over data that was still in the OLD columns (F was labelled
'  "Price (SPY)" but held TWII - that is where 6109% came from), and the
'  migration then refused to start because the header already looked done.
'
'  So this no longer tries to work out which old column held what.  It keeps
'  A (date), B (market value) and C (cumulative PnL) - the three columns the
'  broken run never overwrote - and re-downloads all four index closes for
'  every logged date, which is the authoritative source anyway.
'
'  Realized PnL (D) cannot be recovered on old rows (the old column G was
'  overwritten by a formula); it was 0 on every row, so it is set to 0.
'  N:S is wiped - the old layout and the first baseline block left junk
'  there.  Column Z (GitHub token) is never touched.
' ================================================================
Sub RepairHistoryLog()
    Dim wsH As Worksheet
    On Error Resume Next
    Set wsH = ThisWorkbook.Sheets(SH_HIST)
    On Error GoTo 0
    If wsH Is Nothing Then MsgBox "History sheet not found": Exit Sub

    Dim lastRow As Long
    lastRow = wsH.cells(wsH.Rows.count, "A").End(xlUp).row
    Dim n As Long: n = lastRow - 1
    If n < 1 Then MsgBox "No data rows to repair.", vbInformation: Exit Sub

    If MsgBox("Rebuild HistoryLog index prices from Yahoo history?" & vbLf & vbLf & _
              n & " rows: date / market value / cumulative PnL are kept," & vbLf & _
              "all four index prices are re-downloaded per date," & vbLf & _
              "columns N:S are cleared and Ret% switched to YTD." & vbLf & vbLf & _
              "Realized PnL (col D) is RESET TO 0 on every row - the broken" & vbLf & _
              "run overwrote it, so do not re-run this once real realized" & vbLf & _
              "PnL has been logged." & vbLf & vbLf & _
              "Save a copy of the workbook first.", _
              vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    wsH.Range("N1:S" & lastRow).ClearContents

    Call EnsureHistoryHeaders(wsH)
    Call EnsureYTDBaselines(True)

    Dim tickers As Variant, priceCols As Variant
    tickers = HistTickers()
    priceCols = HistPriceCols()

    Dim missing As Long, r As Long, k As Long
    For r = 2 To lastRow
        Dim targetDate As Date
        targetDate = 0
        On Error Resume Next
        targetDate = CDate(Int(CDbl(wsH.cells(r, "A").Value)))
        On Error GoTo 0

        ' D currently holds whatever the broken run left there (old SPY
        ' prices on the pre-repair rows), and the real values are gone.
        wsH.cells(r, "D").Value = 0
        wsH.Range("B" & r & ":D" & r).NumberFormat = "#,##0"

        For k = LBound(tickers) To UBound(tickers)
            Dim price As Double: price = 0
            If targetDate > 0 Then price = GetHistoricalPrice(CStr(tickers(k)), targetDate)
            If price > 0 Then
                wsH.cells(r, priceCols(k)).Value = price
            Else
                wsH.cells(r, priceCols(k)).ClearContents
                missing = missing + 1
            End If
            wsH.cells(r, priceCols(k)).NumberFormat = "#,##0.00"
        Next k

        Call WriteHistoryRowFormulas(wsH, r)
        Application.StatusBar = "Repairing row " & r & "/" & lastRow
    Next r

    wsH.Range("A:M").EntireColumn.AutoFit
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Repair done - " & n & " rows rebuilt." & vbLf & _
           missing & " index prices could not be downloaded (left blank)." & vbLf & vbLf & _
           "Baselines live on the " & SH_HRAW & " sheet; every Ret% divides" & vbLf & _
           "by its own index's first-trading-day close.", vbInformation
End Sub

' ================================================================
'  Sheet style reset
' ================================================================
Private Sub ResetSheetStyle(ws As Worksheet)
    ws.cells.Clear
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Bold = False
        .NumberFormat = "General"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' A is a blank spacer (as wide as E). B ticker, C name, F sector are the wide text
    ' columns; Q SWING RISK is free text; R/S hold the panel's PnL / return.
    ws.Columns(1).ColumnWidth = 11      ' v4.3: same as E (was 2)
    Dim i As Integer
    For i = 2 To 19
        Select Case i
            Case 2: ws.Columns(i).ColumnWidth = 17
            Case 3, 6: ws.Columns(i).ColumnWidth = 40
            Case 17: ws.Columns(i).ColumnWidth = 20
            Case 18: ws.Columns(i).ColumnWidth = 13
            Case Else: ws.Columns(i).ColumnWidth = 11
        End Select
    Next i
    ws.Columns(RR4_ORD_COL).ColumnWidth = 4
    ws.Columns(RR4_ORD_COL).Hidden = True

    Dim r As Long
    For r = 1 To 200
        ws.Rows(r).RowHeight = 18
    Next r
    ws.Rows(RR4_TOP + 1).RowHeight = 28
End Sub

' ================================================================
'  Header (page rows 1-4 = sheet rows 4-7): total, USD/TWD, ARRANGE <GO>,
'  weight-bar label
' ================================================================
Private Sub DrawHeader(ws As Worksheet, totalMkt As Double, exRate As Double, _
                       arrCode As String)
    With ws.cells(RR4_TOP + 1, RR4_LEFT + 1)
        .Value = totalMkt
        .NumberFormat = "$#,##0"
        .Font.Color = RR4_ACCENT
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ' (v4.4: the " <- TOTAL MKT" caption next to the figure is gone)

    ws.cells(RR4_TOP + 2, RR4_LEFT + 1).Value = "USD/TWD"
    ws.cells(RR4_TOP + 2, RR4_LEFT + 1).Font.Color = RGB(150, 150, 150)
    ws.cells(RR4_TOP + 2, RR4_LEFT + 1).HorizontalAlignment = xlCenter
    With ws.cells(RR4_TOP + 2, RR4_LEFT + 2)
        .Value = exRate
        .NumberFormat = "0.00"
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With

    With ws.cells(RR4_TOP + 2, RR4_LEFT + 3)
        .Value = "ARRANGE"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(RR4_ARR_CELL)
        .NumberFormat = "@"
        .Value = arrCode
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ' E2 (code list + current order) is written by ApplyArrange

    With ws.Range(ws.cells(RR4_TOP + 2, RR4_LEFT + 1), ws.cells(RR4_TOP + 2, RR4_LEFT + 7)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    ws.cells(RR4_TOP + 4, RR4_LEFT + 1).Value = "WEIGHT"
    ws.cells(RR4_TOP + 4, RR4_LEFT + 1).Font.Color = RR4_ACCENT
    ws.cells(RR4_TOP + 4, RR4_LEFT + 1).Font.Bold = True
    ws.cells(RR4_TOP + 4, RR4_LEFT + 1).HorizontalAlignment = xlRight
End Sub

' ================================================================
'  DAILY LOG (page rows 6-12 = sheet rows 9-15) - today only, not kept
' ----------------------------------------------------------------
'  One line per BUY / SELL dated today on Transactions:
'    NEW  = a buy into a ticker that held no shares before today
'    ADD  = a buy into an existing position
'    TRIM = a sell that leaves shares
'    EXIT = a sell that closes the position
'  TODAY line:
'    NET    = cumulative PnL now - HistoryLog col C on the last row dated
'             before today (shows "-" when there is no earlier row)
'    R.PNL  = realized PnL excluding sells dated today => including them
'    P.SIZE = market value minus today's net trade flow => market value
'             (buys add their cost, sells take their proceeds out; TWD)
' ================================================================
Private Sub DrawDailyLog(ws As Worksheet, posData() As Variant, exRate As Double, _
                         totalMkt As Double, totalPnL As Double, realPnL As Double, _
                         prevPnL As Variant)
    With ws.cells(RR4_TOP + 6, RR4_LEFT + 1)
        .Value = "DAILY LOG - " & Format(Date, "yyyy/m/d")
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
    End With

    ' shares held now per ticker (all brokers together)
    Dim held As Object: Set held = CreateObject("Scripting.Dictionary")
    Dim n As Long: n = PosCountOf(posData)
    Dim i As Long
    For i = 1 To n
        Dim hk As String: hk = UCase(CStr(posData(i, 1)))
        held(hk) = held(hk) + posData(i, 9)
    Next i

    Dim wsTr As Worksheet: Set wsTr = ThisWorkbook.Sheets(SH_TRANS)
    Dim lastR As Long: lastR = wsTr.cells(wsTr.Rows.count, "A").End(xlUp).row
    Dim tRows() As Long: ReDim tRows(1 To IIf(lastR > 1, lastR, 1))
    Dim cnt As Long

    ' pass 1: today's rows, and today's net shares per ticker
    Dim todayNet As Object: Set todayNet = CreateObject("Scripting.Dictionary")
    Dim r As Long, act As String, tk As String, sh As Double
    For r = 2 To lastR
        Dim dv As Variant: dv = wsTr.cells(r, "B").Value
        If IsDate(dv) Then
            If Int(CDate(dv)) = Date Then
                act = UCase(CellStr(wsTr.cells(r, "D").Value))
                If act = "BUY" Or act = "SELL" Then
                    cnt = cnt + 1
                    tRows(cnt) = r
                    tk = UCase(CellStr(wsTr.cells(r, "C").Value))
                    sh = NumOr0(wsTr.cells(r, "E").Value)
                    If act = "BUY" Then
                        todayNet(tk) = todayNet(tk) + sh
                    Else
                        todayNet(tk) = todayNet(tk) - sh
                    End If
                End If
            End If
        End If
    Next r

    ' pass 2: one line per trade
    Dim flowTWD As Double, k As Long, outRow As Long: outRow = RR4_TOP + 7
    For k = 1 To cnt
        r = tRows(k)
        act = UCase(CellStr(wsTr.cells(r, "D").Value))
        tk = UCase(CellStr(wsTr.cells(r, "C").Value))
        sh = NumOr0(wsTr.cells(r, "E").Value)
        Dim px As Double: px = NumOr0(wsTr.cells(r, "F").Value)
        Dim amt As Double: amt = NumOr0(wsTr.cells(r, "I").Value)
        Dim fx As Double
        If GetCurrencyType(CStr(tk)) = "USD" Then fx = exRate Else fx = 1
        Dim amtTWD As Double: amtTWD = amt * fx
        If act = "BUY" Then flowTWD = flowTWD + amtTWD Else flowTWD = flowTWD - amtTWD

        Dim nowSh As Double: nowSh = 0
        If held.Exists(tk) Then nowSh = held(tk)
        Dim lineTag As String
        If act = "BUY" Then
            If nowSh - todayNet(tk) <= 0.0001 Then lineTag = "NEW" Else lineTag = "ADD"
        Else
            If nowSh <= 0.0001 Then lineTag = "EXIT" Else lineTag = "TRIM"
        End If

        If cnt <= RR4_LOG_ROWS Or k < RR4_LOG_ROWS Then
            Call WriteLogLine(ws, outRow, lineTag, tk, sh, px, amtTWD)
            outRow = outRow + 1
        End If
    Next k

    If cnt = 0 Then
        With ws.cells(RR4_TOP + 7, RR4_LEFT + 2)
            .Value = "NO TRADES TODAY"
            .Font.Color = RGB(150, 150, 150)
            .Font.Italic = True
        End With
    ElseIf cnt > RR4_LOG_ROWS Then
        With ws.cells(outRow, RR4_LEFT + 2)
            .Value = "+" & (cnt - RR4_LOG_ROWS + 1) & " more trades today - see Transactions"
            .Font.Color = RGB(150, 150, 150)
            .Font.Italic = True
        End With
    End If

    ' --- TODAY line ---
    Dim realBefore As Double: realBefore = RealizedBefore(Date)
    Dim sizeBefore As Double: sizeBefore = totalMkt - flowTWD
    Dim netTxt As String, netVal As Double
    If IsEmpty(prevPnL) Then
        netTxt = "-"
    Else
        netVal = totalPnL - CDbl(prevPnL)
        netTxt = Format(netVal, "+#,##0;-#,##0;0")
    End If

    With ws.cells(RR4_TOP + 12, RR4_LEFT + 1)
        .Value = "TODAY"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
    End With
    With ws.cells(RR4_TOP + 12, RR4_LEFT + 2)
        .Value = "NET " & netTxt & "   |   R.PNL " & Format(realBefore, "#,##0") & _
                 " => " & Format(realPnL, "#,##0") & "   |   P.SIZE " & _
                 Format(sizeBefore, "#,##0") & " => " & Format(totalMkt, "#,##0")
        .Font.Color = RGB(221, 221, 221)
        If Not IsEmpty(prevPnL) Then
            .Characters(5, Len(netTxt)).Font.Color = GainLossColor(netVal)
            .Characters(5, Len(netTxt)).Font.Bold = True
        End If
    End With
    With ws.Range(ws.cells(RR4_TOP + 12, RR4_LEFT + 1), ws.cells(RR4_TOP + 12, RR4_LEFT + 7)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With
End Sub

Private Sub WriteLogLine(ws As Worksheet, r As Long, lineTag As String, tk As String, _
                         sh As Double, px As Double, amtTWD As Double)
    With ws.cells(r, RR4_LEFT + 1)
        .Value = lineTag
        .Font.Bold = True
        If lineTag = "NEW" Or lineTag = "ADD" Then
            .Font.Color = RR4_ACCENT
        Else
            .Font.Color = RGB(0, 200, 255)
        End If
    End With
    With ws.cells(r, RR4_LEFT + 2)
        .Value = tk & "   " & FormatShares(sh) & " sh   @ " & Format(px, "#,##0.00") & _
                 "   |   " & Format(amtTWD, "#,##0") & " TWD"
        .Font.Color = RGB(221, 221, 221)
    End With
End Sub

' ================================================================
'  SUMMARY (page rows 14-20 = sheet rows 17-23): labels B / F, values C / G
'  Row 20 (v4.6) = CURRENCY  USD nn% : TWD nn%, by market value in TWD -
'  the one thing kept from the deleted Analysis page.
' ================================================================
Private Sub DrawSummary(ws As Worksheet, posData() As Variant, totalMkt As Double, _
                        totalCost As Double, totalUnrl As Double, realPnL As Double, _
                        portBeta As Double, posCount As Long)
    With ws.cells(RR4_TOP + 14, RR4_LEFT + 1)
        .Value = "SUMMARY - RR4"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
    End With

    Dim inc As Date: inc = GetInceptionDate(ws)
    Dim startCap As Double: startCap = GetStartingCapital(ws)
    Dim rlPct As Double, unrlPct As Double
    If startCap > 0 Then rlPct = realPnL / startCap
    If totalCost > 0 Then unrlPct = totalUnrl / totalCost

    Call SumKV(ws, RR4_TOP + 15, RR4_LEFT + 1, "INCEPTION", inc, "yyyy-mm-dd", False)
    Call SumKV(ws, RR4_TOP + 16, RR4_LEFT + 1, "DAYS", CDbl(Date - inc), "0", False)
    Call SumKV(ws, RR4_TOP + 17, RR4_LEFT + 1, "POSITIONS", CDbl(posCount), "0", False)
    Call SumKV(ws, RR4_TOP + 18, RR4_LEFT + 1, "REALISED PNL", realPnL, "+#,##0;-#,##0;0", True)
    Call SumKV(ws, RR4_TOP + 19, RR4_LEFT + 1, "UNREALISED PNL", totalUnrl, "+#,##0;-#,##0;0", True)

    Call SumKV(ws, RR4_TOP + 15, RR4_LEFT + 5, "STARTING", startCap, "#,##0", False)
    Call SumKV(ws, RR4_TOP + 16, RR4_LEFT + 5, "NET EXPOSURE", totalMkt, "#,##0", False)
    Call SumKV(ws, RR4_TOP + 17, RR4_LEFT + 5, "PORT.BETA", portBeta, "0.000", False)
    ws.cells(RR4_TOP + 17, RR4_LEFT + 6).Font.Color = RR4_ACCENT
    Call SumKV(ws, RR4_TOP + 18, RR4_LEFT + 5, "REALISED PNL %", rlPct, "+0.00%;-0.00%;0.00%", True)
    Call SumKV(ws, RR4_TOP + 19, RR4_LEFT + 5, "UNREALISED PNL %", unrlPct, "+0.00%;-0.00%;0.00%", True)

    ' currency allocation (market value, TWD terms)
    Dim usdMV As Double, twdMV As Double, n As Long, i As Long
    n = PosCountOf(posData)
    For i = 1 To n
        If GetCurrencyType(CStr(posData(i, 1))) = "USD" Then
            usdMV = usdMV + posData(i, 8)
        Else
            twdMV = twdMV + posData(i, 8)
        End If
    Next i
    Dim curTxt As String
    If usdMV + twdMV > 0 Then
        curTxt = "USD " & Format(usdMV / (usdMV + twdMV), "0%") & " : TWD " & Format(twdMV / (usdMV + twdMV), "0%")
    Else
        curTxt = "-"
    End If
    Call SumKV(ws, RR4_TOP + 20, RR4_LEFT + 1, "CURRENCY", curTxt, "@", False)
End Sub

Private Sub SumKV(ws As Worksheet, r As Long, c As Long, lbl As String, _
                  ByVal v As Variant, fmt As String, colorPnL As Boolean)
    With ws.cells(r, c)
        .Value = lbl
        .Font.Color = RGB(0, 200, 255)
    End With
    With ws.cells(r, c + 1)
        .NumberFormat = fmt
        .Value = v
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        If colorPnL Then
            .Font.Color = GainLossColor(CDbl(v))
        Else
            .Font.Color = RGB(221, 221, 221)
        End If
    End With
End Sub

' ================================================================
'  Position log title (RR4_POS_TITLE) + column header row (RR4_POS_HDR)
' ================================================================
Private Sub DrawColumnHeaders(ws As Worksheet)
    With ws.cells(RR4_POS_TITLE, RR4_LEFT + 1)
        .Value = "POSITION LOG - RR4"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
    End With
    With ws.Range(ws.cells(RR4_POS_TITLE, RR4_LEFT + 1), ws.cells(RR4_POS_TITLE, RR4_LEFT + 18)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    Dim headers As Variant
    headers = Array("TICKER", "NAME", "ENTRY DT", "DAYS", "SECTOR", _
                    "NET EXPOS", "SHARES", "ENTRY PX", "LAST", "% CHG", _
                    "UNRL PNL", "WT%", "W.BETA", "BETA 180D", "P.TARGET", "SWING RISK")
    Dim i As Integer
    For i = 0 To UBound(headers)
        With ws.cells(RR4_POS_HDR, RR4_LEFT + i + 1)
            .Value = headers(i)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .Font.Size = 9
            .Font.Name = "Consolas"
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next i
    ws.Rows(RR4_POS_HDR).RowHeight = 20
    With ws.Range(ws.cells(RR4_POS_HDR, RR4_LEFT + 1), ws.cells(RR4_POS_HDR, RR4_NCOL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With
End Sub

' ------------------------------------------------------------
' One row per position, no broker grouping (v4). Rows come out in
' CalcPositions order; ApplyArrange then sorts them and paints the
' stripes, the P.TARGET highlight and the closing gold rule.
Private Function WritePositionRows(ws As Worksheet, posData() As Variant, _
                                    totalMkt As Double, ByVal swingRiskMap As Object) As Long
    WritePositionRows = RR4_POS_FIRST - 1
    Dim n As Long: n = PosCountOf(posData)
    If n < 1 Then Exit Function

    Dim i As Long, r As Long: r = RR4_POS_FIRST
    For i = 1 To n
        Call WriteOnePositionRow(ws, r, i, posData, totalMkt, swingRiskMap)
        r = r + 1
    Next i
    WritePositionRows = r - 1

    Call DrawTopExposure(ws, posData, totalMkt)
End Function

' ------------------------------------------------------------
Private Sub WriteOnePositionRow(ws As Worksheet, r As Long, i As Long, _
                                 posData() As Variant, totalMkt As Double, _
                                 swingRiskMap As Object)
    With ws.Range(ws.cells(r, RR4_LEFT + 1), ws.cells(r, RR4_NCOL))
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Color = RGB(210, 210, 210)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = RR4_POS_ROW_H

    Dim tickerCode As String: tickerCode = CStr(posData(i, 1))
    Dim nm         As String: nm = posData(i, 3)
    Dim entryDt    As Date:   entryDt = posData(i, 4)
    Dim days       As Long:   days = posData(i, 5)
    Dim Sector     As String: Sector = posData(i, 7)
    Dim netExpos   As Double: netExpos = posData(i, 8)
    Dim shares     As Double: shares = posData(i, 9)
    Dim entryPx    As Double: entryPx = posData(i, 10)
    Dim lastPx     As Double: lastPx = posData(i, 11)
    Dim chgPct     As Double: chgPct = posData(i, 12)
    Dim unrlPnl    As Double: unrlPnl = posData(i, 13)
    Dim Beta       As Double: Beta = posData(i, 14)
    Dim beta30d    As Double: beta30d = posData(i, 15)
    Dim pTgt       As Double: pTgt = posData(i, 16)

    Dim wtPct As Double: If totalMkt > 0 Then wtPct = netExpos / totalMkt
    Dim wBeta As Double: wBeta = wtPct * Beta

    ws.cells(r, RR4_LEFT + 1).Value = tickerCode
    ws.cells(r, RR4_LEFT + 1).Font.Color = RR4_ACCENT
    ws.cells(r, RR4_LEFT + 1).Font.Bold = True

    ws.cells(r, RR4_LEFT + 2).Value = nm
    ws.cells(r, RR4_LEFT + 2).HorizontalAlignment = xlLeft
    ws.cells(r, RR4_LEFT + 2).Font.Color = RGB(221, 221, 221)

    ws.cells(r, RR4_LEFT + 3).Value = entryDt
    ws.cells(r, RR4_LEFT + 3).NumberFormat = "m/d/yyyy"

    ws.cells(r, RR4_LEFT + 4).Value = days

    ws.cells(r, RR4_LEFT + 5).Value = Sector
    ws.cells(r, RR4_LEFT + 5).Font.Color = RGB(180, 180, 180)

    ws.cells(r, RR4_LEFT + 6).Value = netExpos
    ws.cells(r, RR4_LEFT + 6).NumberFormat = "$#,##0"

    ws.cells(r, RR4_LEFT + 7).Value = shares
    ws.cells(r, RR4_LEFT + 7).NumberFormat = "General"

    ws.cells(r, RR4_LEFT + 8).Value = entryPx
    ws.cells(r, RR4_LEFT + 8).NumberFormat = "#,##0.00"

    ws.cells(r, RR4_LEFT + 9).Value = lastPx
    ws.cells(r, RR4_LEFT + 9).NumberFormat = "#,##0.00"
    ws.cells(r, RR4_LEFT + 9).Font.Color = RGB(221, 221, 221)

    ws.cells(r, RR4_LEFT + 10).Value = chgPct
    ws.cells(r, RR4_LEFT + 10).NumberFormat = "0.00%"
    ws.cells(r, RR4_LEFT + 10).Font.Bold = True
    ws.cells(r, RR4_LEFT + 10).Font.Color = PnLColorMuted(Round(unrlPnl, 0))

    ws.cells(r, RR4_LEFT + 11).Value = unrlPnl
    ws.cells(r, RR4_LEFT + 11).NumberFormat = "$#,##0;-$#,##0"
    ws.cells(r, RR4_LEFT + 11).Font.Bold = True
    ws.cells(r, RR4_LEFT + 11).Font.Color = PnLColorMuted(Round(unrlPnl, 0))

    ws.cells(r, RR4_LEFT + 12).Value = wtPct
    ws.cells(r, RR4_LEFT + 12).NumberFormat = "0.00%"

    ws.cells(r, RR4_LEFT + 13).Value = wBeta
    ws.cells(r, RR4_LEFT + 13).NumberFormat = "0.000"

    ws.cells(r, RR4_LEFT + 14).Value = beta30d
    ws.cells(r, RR4_LEFT + 14).NumberFormat = "0.000"
    ws.cells(r, RR4_LEFT + 14).Font.Color = RGB(180, 180, 255)

    ws.cells(r, RR4_LEFT + 15).Value = pTgt
    ws.cells(r, RR4_LEFT + 15).NumberFormat = "#,##0"
    ws.cells(r, RR4_LEFT + 15).Font.Color = RR4_ACCENT

    If swingRiskMap.Exists(tickerCode) Then
        With ws.cells(r, RR4_SWING_COL)
            .Value = swingRiskMap(tickerCode)
            .Font.Color = RR4_INPUT_FG
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End If

    ' hidden default-order key (ARRANGE "DEF")
    ws.cells(r, RR4_ORD_COL).NumberFormat = "@"
    ws.cells(r, RR4_ORD_COL).Value = DefaultOrderKey(tickerCode)
End Sub

' TW codes first, numerically (2330 before 3017), then US tickers A-Z.
Private Function DefaultOrderKey(ByVal t As String) As String
    If GetCurrencyType(CStr(t)) = "TWD" Then
        DefaultOrderKey = "0" & Format(val(t), "0000000000") & UCase(t)
    Else
        DefaultOrderKey = "1" & UCase(t)
    End If
End Function

' "TOP EXPOSURE  2330 6.1%  NVDA 5.7% ..." next to the position-log title
Private Sub DrawTopExposure(ws As Worksheet, posData() As Variant, totalMkt As Double)
    Dim n As Long: n = PosCountOf(posData)
    If n < 1 Or totalMkt <= 0 Then Exit Sub

    Dim used() As Boolean: ReDim used(1 To n)
    Dim s As String, k As Long, i As Long, best As Long
    For k = 1 To IIf(n < 5, n, 5)
        best = 0
        For i = 1 To n
            If Not used(i) Then
                If best = 0 Then
                    best = i
                ElseIf posData(i, 8) > posData(best, 8) Then
                    best = i
                End If
            End If
        Next i
        used(best) = True
        s = s & ShortTicker(CStr(posData(best, 1))) & " " & _
            Format(posData(best, 8) / totalMkt, "0.0%") & "    "
    Next k

    With ws.cells(RR4_POS_TITLE, RR4_LEFT + 5)
        .Value = "TOP EXPOSURE    " & s
        .Font.Color = RGB(180, 180, 180)
        .Characters(1, 12).Font.Color = RR4_ACCENT
        .Characters(1, 12).Font.Bold = True
    End With
End Sub

' ================================================================
'  ARRANGE <GO> - sort the position rows already on the sheet
' ----------------------------------------------------------------
'  Called by RebuildPortfolioDashboard and by the RR4 sheet's
'  Worksheet_Change when D2 is edited. Uses Range.Sort, so values AND
'  their per-cell fonts move together; the stripes / P.TARGET highlight
'  / closing rule depend on row position and are repainted afterwards.
'  Key 2 is always the hidden default-order key, so ties stay stable.
' ================================================================
Public Sub ApplyArrange(Optional ByVal code As String = vbNullString)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_PORT)
    code = UCase(Trim(code))

    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Fin

    Dim keyCol As Long, ord As Long, desc As String, known As Boolean
    known = True
    Select Case code
        Case "", "DEF": keyCol = RR4_ORD_COL: ord = xlAscending: desc = "DEFAULT (TW code, then US A-Z)"
        Case "UNU": keyCol = 11: ord = xlAscending: desc = "UNRL PNL  low > high"
        Case "UND": keyCol = 11: ord = xlDescending: desc = "UNRL PNL  high > low"
        Case "PCU": keyCol = 10: ord = xlAscending: desc = "% CHG  low > high"
        Case "PCD": keyCol = 10: ord = xlDescending: desc = "% CHG  high > low"
        Case "DAU": keyCol = 4: ord = xlAscending: desc = "DAYS  short > long"
        Case "DAD": keyCol = 4: ord = xlDescending: desc = "DAYS  long > short"
        Case "WTU": keyCol = 12: ord = xlAscending: desc = "WT%  light > heavy"
        Case "WTD": keyCol = 12: ord = xlDescending: desc = "WT%  heavy > light"
        Case Else
            known = False
            keyCol = RR4_ORD_COL: ord = xlAscending
            desc = "UNKNOWN CODE [" & code & "] - default order"
    End Select

    Dim lastR As Long: lastR = LastPositionRow(ws)
    If lastR > RR4_POS_FIRST Then
        Dim rng As Range
        Set rng = ws.Range(ws.cells(RR4_POS_FIRST, RR4_LEFT + 1), ws.cells(lastR, RR4_ORD_COL))
        ' The hidden default-order column is the tie-breaker - but ONLY when it
        ' is not already Key1. Passing the same column as Key1 and Key2 makes
        ' Excel write <sortCondition ref="T.."/> TWICE into the sheet's
        ' sortState; it saves without complaint and then REFUSES TO OPEN the
        ' file ever again (error 800A03EC, no repair offered). That is what
        ' broke the workbook on 2026-09-12 - every "DEF" sort hit it.
        If keyCol = RR4_ORD_COL Then
            rng.Sort Key1:=ws.cells(RR4_POS_FIRST, keyCol), Order1:=ord, _
                     Header:=xlNo, Orientation:=xlTopToBottom, MatchCase:=False
        Else
            rng.Sort Key1:=ws.cells(RR4_POS_FIRST, keyCol), Order1:=ord, _
                     Key2:=ws.cells(RR4_POS_FIRST, RR4_ORD_COL), Order2:=xlAscending, _
                     Header:=xlNo, Orientation:=xlTopToBottom, MatchCase:=False
        End If
    End If
    If lastR >= RR4_POS_FIRST Then Call RestripeRows(ws, lastR)
    Call DrawWeightBar(ws, lastR)
    Call DrawDonut(ws, lastR)

    Dim codeList As String: codeList = "UNU/UND PCU/PCD DAU/DAD WTU/WTD DEF  >>  "
    With ws.cells(RR4_TOP + 2, RR4_LEFT + 5)
        .Value = codeList & desc
        .Font.Size = 9
        .Font.Color = RGB(150, 150, 150)
        .Characters(Len(codeList) + 1, Len(desc)).Font.Color = IIf(known, RR4_ACCENT, RGB(255, 80, 80))
    End With

Fin:
    ' never leave the screen frozen behind an error - the sheet's own handler
    ' only restores EnableEvents
    Dim errN As Long: errN = Err.Number
    Application.ScreenUpdating = prevScr
    If errN <> 0 Then Call NavNotify("ARRANGE failed: " & Err.Description, True)
End Sub

Private Function IsArrangeCode(ByVal code As String) As Boolean
    Select Case UCase(Trim(code))
        Case "", "DEF", "UNU", "UND", "PCU", "PCD", "DAU", "DAD", "WTU", "WTD"
            IsArrangeCode = True
    End Select
End Function

' Last position row = last consecutive row carrying a default-order key.
' Bounded by the last used cell of column T: an unbounded walk would run off
' the sheet (error 1004) if that column were ever filled to the bottom, and a
' stale key below the block would stretch the sort range into the disclaimer's
' merged rows ("merged cells must be identically sized").
Private Function LastPositionRow(ws As Worksheet) As Long
    Dim lastUsed As Long
    lastUsed = ws.cells(ws.Rows.count, RR4_ORD_COL).End(xlUp).row
    Dim r As Long: r = RR4_POS_FIRST
    Do While r <= lastUsed
        If CellStr(ws.cells(r, RR4_ORD_COL).Value) = "" Then Exit Do
        r = r + 1
    Loop
    LastPositionRow = r - 1
End Function

Private Sub RestripeRows(ws As Worksheet, lastR As Long)
    Dim r As Long, bg As Long
    For r = RR4_POS_FIRST To lastR
        ' v4.4: one step darker than the 15 / 22 used on the other pages
        If (r - RR4_POS_FIRST) Mod 2 = 0 Then bg = RGB(8, 8, 8) Else bg = RGB(14, 14, 14)
        Dim lastPx As Double: lastPx = NumOr0(ws.cells(r, RR4_LEFT + 9).Value)
        Dim pTgt As Double: pTgt = NumOr0(ws.cells(r, RR4_LEFT + 15).Value)
        If pTgt > 0 And lastPx > pTgt Then bg = RGB(40, 25, 0)
        With ws.Range(ws.cells(r, RR4_LEFT + 1), ws.cells(r, RR4_NCOL))
            .Interior.Color = bg
            .Borders(xlEdgeBottom).LineStyle = xlNone
        End With
        ws.cells(r, RR4_SWING_COL).Interior.Color = RR4_SWING_BG   ' typed by hand
        ws.cells(r, RR4_SWING_COL).Font.Color = RR4_INPUT_FG
    Next r
    With ws.Range(ws.cells(lastR, RR4_LEFT + 1), ws.cells(lastR, RR4_NCOL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With
End Sub

' ================================================================
'  WEIGHT BAR (row 4, B:G) - one rectangle per position, left to right
'  in the current ARRANGE order. Width = WT%, colour = % CHG - TW
'  convention since v4.3: RED up, GREEN down, stronger the further from
'  0, full strength at +/-30%. Ticker text is white.
'  Shapes are named RR4W_n so only they are cleared on a redraw.
' ================================================================
Private Sub DrawWeightBar(ws As Worksheet, lastR As Long)
    Dim k As Long
    For k = ws.Shapes.count To 1 Step -1
        If Left(ws.Shapes(k).Name, Len(RR4_WBAR_PREFIX)) = RR4_WBAR_PREFIX Then ws.Shapes(k).Delete
    Next k
    If lastR < RR4_POS_FIRST Then Exit Sub

    Dim barL As Double, barT As Double, barW As Double, barH As Double
    barL = ws.cells(RR4_TOP + 4, RR4_LEFT + 2).Left
    barW = ws.cells(RR4_TOP + 4, RR4_LEFT + 8).Left - barL
    barT = ws.cells(RR4_TOP + 4, RR4_LEFT + 2).Top + 2
    barH = ws.Rows(RR4_TOP + 4).RowHeight - 4

    Dim sumW As Double, r As Long
    For r = RR4_POS_FIRST To lastR
        sumW = sumW + Abs(NumOr0(ws.cells(r, RR4_LEFT + 12).Value))
    Next r
    If sumW <= 0 Then Exit Sub

    Dim x As Double: x = barL
    For r = RR4_POS_FIRST To lastR
        Dim wt As Double: wt = Abs(NumOr0(ws.cells(r, RR4_LEFT + 12).Value))
        Dim w As Double: w = wt / sumW * barW
        If w >= 1 Then
            Dim pct As Double: pct = NumOr0(ws.cells(r, RR4_LEFT + 10).Value)
            Dim tk As String: tk = CellStr(ws.cells(r, RR4_LEFT + 1).Value)
            Dim shp As Shape
            Set shp = ws.Shapes.AddShape(msoShapeRectangle, x, barT, IIf(w > 2, w - 1, w), barH)
            shp.Name = RR4_WBAR_PREFIX & (r - RR4_POS_FIRST + 1)
            shp.Line.Visible = msoFalse
            shp.Fill.ForeColor.RGB = WeightBarColor(pct)
            shp.Placement = xlMove
            shp.AlternativeText = tk & "  WT " & Format(wt, "0.0%") & "  " & Format(pct, "+0.0%;-0.0%;0.0%")
            If w >= 42 Then
                With shp.TextFrame2
                    .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                    .WordWrap = msoFalse
                    .VerticalAnchor = msoAnchorMiddle
                    .TextRange.Text = ShortTicker(tk)
                    .TextRange.Font.Name = "Consolas"
                    .TextRange.Font.Size = 7
                    .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                End With
            End If
        End If
        x = x + w
    Next r
End Sub

Private Function WeightBarColor(ByVal pct As Double) As Long
    Dim t As Double: t = Abs(pct) / 0.3
    If t > 1 Then t = 1
    t = 0.35 + 0.65 * t
    Dim r1 As Double, g1 As Double, b1 As Double
    ' red = up, green = down (TW convention, v4.3)
    If pct >= 0 Then
        r1 = 235: g1 = 70: b1 = 70
    Else
        r1 = 0: g1 = 200: b1 = 90
    End If
    ' blend from dark grey (45,45,45) towards the full colour
    WeightBarColor = RGB(CLng(45 + (r1 - 45) * t), CLng(45 + (g1 - 45) * t), CLng(45 + (b1 - 45) * t))
End Function

' v4.1: the eight ActiveX buttons (DELETE ALL .. VOLATILITY) are replaced
' by typed nav commands (UP / ADD / DEL / V! / D! / C! / DBG / CLEARALL,
' see modNav). Any CmdBtn* still on the sheet is removed on the next UP.
Private Sub RemoveLegacyButtons(ws As Worksheet)
    Dim k As Long
    On Error Resume Next
    For k = ws.OLEObjects.count To 1 Step -1
        If Left(ws.OLEObjects(k).Name, 6) = "CmdBtn" Then ws.OLEObjects(k).Delete
    Next k
    On Error GoTo 0
End Sub

' ================================================================
'  WEIGHT DONUT (right of the ticker panel, from column U)
'  Same data as the weight bar: one slice per position in the current
'  ARRANGE order, size = WT%. Colours come from a fixed palette
'  (DonutColor, one per slice - NOT the up/down colour of the bar, v4.3),
'  no data labels. The series
'  points at the position-log cells (tickers A, WT% L) instead of holding
'  copied numbers, so what the chart shows can always be checked on the
'  sheet. Rebuilt with the weight bar (UP and every ARRANGE).
' ================================================================
Private Sub DrawDonut(ws As Worksheet, lastR As Long)
    ' delete by walking the collection - ChartObjects("name") raises when it is
    ' missing, and an ignored error there would leave a second chart behind
    Dim k As Long
    For k = ws.ChartObjects.count To 1 Step -1
        If ws.ChartObjects(k).Name = RR4_DONUT_NAME Then ws.ChartObjects(k).Delete
    Next k
    If lastR < RR4_POS_FIRST Then Exit Sub

    ' chart band (v4.7): left edge on column F, full band height
    Dim cL As Double, cT As Double, sz As Double
    cL = ws.Columns(RR4_LEFT + 5).Left + 4
    cT = ws.Rows(RR4_CHART_TOP).Top + 4
    sz = ws.Rows(RR4_CHART_TOP + RR4_CHART_ROWS).Top - cT - 4
    ' ApplyArrange also runs from the E5 edit, where nothing has re-set the row
    ' heights: a hidden or squashed block would ask for a zero-height chart
    If sz < 60 Then sz = 60

    Dim co As ChartObject
    ' wider than tall: doughnut labels can only sit ON the ring (Excel has
    ' no outside position for doughnuts), so the ring is kept thick
    Set co = ws.ChartObjects.Add(cL, cT, sz * 1.5, sz)
    co.Name = RR4_DONUT_NAME
    co.Placement = xlMove
    With co.Chart
        ' v4.10: a PIE, not a doughnut - Excel can only put doughnut labels
        ' on the ring, and the labels are wanted outside with leader lines.
        ' The hole is faked with a black circle drawn over the centre below.
        .ChartType = xlPie
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        Dim ser As Series
        Set ser = .SeriesCollection.NewSeries
        ser.Name = "WT%"
        ser.Values = ws.Range(ws.cells(RR4_POS_FIRST, RR4_LEFT + 12), ws.cells(lastR, RR4_LEFT + 12))
        ser.XValues = ws.Range(ws.cells(RR4_POS_FIRST, RR4_LEFT + 1), ws.cells(lastR, RR4_LEFT + 1))
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = "WEIGHT"
        .ChartTitle.Font.Name = "Consolas"
        .ChartTitle.Font.Size = 9
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RR4_ACCENT
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)

        Dim r As Long, p As Long
        For r = RR4_POS_FIRST To lastR
            p = r - RR4_POS_FIRST + 1
            With ser.Points(p).Format
                .Fill.ForeColor.RGB = DonutColor(p)
                .Line.ForeColor.RGB = RGB(0, 0, 0)
                .Line.Weight = 1.5
            End With
        Next r

        ' "3653.TW 14.9%" outside each slice with a leader line (v4.10)
        ser.HasDataLabels = True
        With ser.DataLabels
            .ShowSeriesName = False
            .ShowValue = False
            .ShowLegendKey = False
            .ShowCategoryName = True
            .ShowPercentage = True
            .Separator = " "
            .NumberFormat = "0.0%"
            .Position = xlLabelPositionOutsideEnd
            .Font.Name = "Consolas"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With
        ser.HasLeaderLines = True
        On Error Resume Next
        ser.LeaderLines.Format.Line.ForeColor.RGB = RGB(120, 120, 120)
        ser.LeaderLines.Format.Line.Weight = 0.75
        On Error GoTo 0

        ' fake the 56% hole: a black circle centred on the plot area
        Dim pa As PlotArea: Set pa = .PlotArea
        Dim d As Double: d = IIf(pa.InsideWidth < pa.InsideHeight, pa.InsideWidth, pa.InsideHeight) * 0.56
        Dim hole As Shape
        Set hole = .Shapes.AddShape(msoShapeOval, _
                     pa.InsideLeft + (pa.InsideWidth - d) / 2, _
                     pa.InsideTop + (pa.InsideHeight - d) / 2, d, d)
        hole.Name = "RR4_DONUT_HOLE"
        hole.Fill.ForeColor.RGB = RGB(0, 0, 0)
        hole.Line.Visible = msoFalse
    End With
End Sub

' ================================================================
'  REALIZED PNL LINE (v4.6) - under the weight donut, same width.
'  One series: HistoryLog column D (cumulative realized PnL, TWD) against
'  column A (date), rows 2..last. Points at the sheet cells, so the line
'  follows whatever LogHistory / BackfillHistory write there.
' ================================================================
Private Sub DrawRealizedChart(ws As Worksheet)
    Dim k As Long
    For k = ws.ChartObjects.count To 1 Step -1
        If ws.ChartObjects(k).Name = RR4_RLPNL_NAME Then ws.ChartObjects(k).Delete
    Next k

    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = ThisWorkbook.Sheets(SH_HIST): On Error GoTo 0
    If wsH Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = wsH.cells(wsH.Rows.count, "A").End(xlUp).row
    If lastR < 3 Then Exit Sub          ' one point is not a line

    ' chart band (v4.7): columns J:S, full band height, next to the donut
    Dim cL As Double, cT As Double, cW As Double, cH As Double
    cL = ws.Columns(10).Left + 4
    cW = ws.Columns(20).Left - cL - 4
    cT = ws.Rows(RR4_CHART_TOP).Top + 4
    cH = ws.Rows(RR4_CHART_TOP + RR4_CHART_ROWS).Top - cT - 4
    If cH < 120 Then cH = 120

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(cL, cT, cW, cH)
    co.Name = RR4_RLPNL_NAME
    co.Placement = xlMove
    With co.Chart
        .ChartType = xlLine
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        Dim ser As Series
        Set ser = .SeriesCollection.NewSeries
        ser.Name = "REALIZED PNL"
        ser.Values = wsH.Range(wsH.cells(2, 4), wsH.cells(lastR, 4))
        ser.XValues = wsH.Range(wsH.cells(2, 1), wsH.cells(lastR, 1))
        ser.Format.Line.ForeColor.RGB = RGB(255, 192, 0)    ' yellow line (v4.6.1)
        ser.Format.Line.Weight = 0.75                       ' 1 px
        ser.MarkerStyle = xlMarkerStyleNone
        ser.Smooth = False

        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = "REALIZED PNL"
        .ChartTitle.Font.Name = "Consolas"
        .ChartTitle.Font.Size = 9
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(255, 255, 255)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)

        With .Axes(xlCategory)
            .CategoryType = xlTimeScale
            .MajorUnitScale = xlMonths
            .MajorUnit = 2                  ' one tick every two months
            .TickLabels.NumberFormat = "yyyy/m"
            .TickLabels.Font.Name = "Consolas"
            .TickLabels.Font.Size = 7
            .TickLabels.Font.Color = RGB(255, 255, 255)
            .Format.Line.ForeColor.RGB = RR4_LINE
            .MajorGridlines.Delete
        End With
        With .Axes(xlValue)
            .TickLabels.NumberFormat = "#,##0"
            .TickLabels.Font.Name = "Consolas"
            .TickLabels.Font.Size = 7
            .TickLabels.Font.Color = RGB(255, 255, 255)
            .Format.Line.Visible = msoFalse
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(30, 30, 30)
        End With
    End With
End Sub

' ================================================================
'  WATCHLIST (B25:E38) - see the RR4_WL_* constants
' ================================================================
' Saved rows, read BEFORE the sheet is cleared. Only trusted when the
' title is in place (an older layout has other things in those cells).
' Returns a 2-D Variant(1..n, 1..3) = ticker / strategy / target, blanks
' compacted out; Empty when there is nothing.
Private Function ReadWatchlist(ws As Worksheet) As Variant
    If Left(UCase(CellStr(ws.cells(RR4_WL_TITLE, RR4_LEFT + 1).Value)), 9) <> "WATCHLIST" Then Exit Function
    Dim tmp(1 To 12, 1 To 3) As Variant, n As Long, r As Long
    For r = RR4_WL_FIRST To RR4_WL_LAST
        Dim tk As String: tk = UCase(CellStr(ws.cells(r, RR4_LEFT + 1).Value))
        Dim st As String: st = CellStr(ws.cells(r, RR4_LEFT + 2).Value)
        Dim tg As Variant: tg = ws.cells(r, RR4_LEFT + 3).Value
        If tk <> "" Then
            n = n + 1
            tmp(n, 1) = tk
            tmp(n, 2) = st
            tmp(n, 3) = IIf(IsNumeric(tg) And Not IsEmpty(tg), CDbl(tg), Empty)
        End If
    Next r
    If n = 0 Then Exit Function
    Dim out() As Variant: ReDim out(1 To n, 1 To 3)
    For r = 1 To n
        out(r, 1) = tmp(r, 1): out(r, 2) = tmp(r, 2): out(r, 3) = tmp(r, 3)
    Next r
    ReadWatchlist = out
End Function

' Title, header, the entry row, then the saved rows with live price.
Private Sub DrawWatchlist(ws As Worksheet, wl As Variant)
    With ws.cells(RR4_WL_TITLE, RR4_LEFT + 1)
        .Value = "WATCHLIST     type in row " & RR4_WL_ENTRY & " + Enter to add  .  double-click a row to delete"
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
        .Font.Bold = False
        .Characters(1, 9).Font.Color = RR4_ACCENT
        .Characters(1, 9).Font.Bold = True
        .Characters(1, 9).Font.Size = 10
    End With
    Dim hdr As Variant: hdr = Array("TICKER", "STRATEGY", "ENTRY TGT", "LAST")
    Dim c As Long
    For c = 0 To 3
        With ws.cells(RR4_WL_HDR, RR4_LEFT + 1 + c)
            .Value = hdr(c)
            .Font.Color = RGB(0, 200, 255)
            .Font.Size = 9
            .HorizontalAlignment = IIf(c = 1, xlLeft, xlCenter)
        End With
    Next c
    With ws.Range(ws.cells(RR4_WL_HDR, RR4_LEFT + 1), ws.cells(RR4_WL_HDR, RR4_LEFT + 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    Call WatchlistPaintEntryRow(ws)

    Dim n As Long: If IsArray(wl) Then n = UBound(wl, 1)
    Dim r As Long, i As Long
    For r = RR4_WL_FIRST To RR4_WL_LAST
        i = r - RR4_WL_FIRST + 1
        If i <= n Then
            ws.cells(r, RR4_LEFT + 1).Value = wl(i, 1)
            ws.cells(r, RR4_LEFT + 2).Value = wl(i, 2)
            If Not IsEmpty(wl(i, 3)) Then ws.cells(r, RR4_LEFT + 3).Value = wl(i, 3)
        End If
        Call RefreshWatchlistRow(ws, r)
    Next r
End Sub

' The entry row: three dark input cells, always empty after a commit.
Private Sub WatchlistPaintEntryRow(ws As Worksheet)
    Dim c As Long
    For c = 1 To 3
        With ws.cells(RR4_WL_ENTRY, RR4_LEFT + c)
            .Value = ""
            .Interior.Color = RR4_INPUT_BG
            .Font.Color = RR4_INPUT_FG
            .Font.Bold = (c = 1)
            .HorizontalAlignment = IIf(c = 2, xlLeft, xlCenter)
            .NumberFormat = IIf(c = 3, "#,##0.00", "@")
        End With
    Next c
    With ws.cells(RR4_WL_ENTRY, RR4_LEFT + 4)
        .Value = ""
        .Interior.Color = RGB(0, 0, 0)
    End With
End Sub

' Called by the sheet code when B/C/D of the entry row changes. Commits
' once ticker AND target are both in; otherwise leaves the row alone so
' the other cells can still be typed.
Public Sub WatchlistCommitEntry(ws As Worksheet)
    Dim tk As String: tk = UCase(Trim(CellStr(ws.cells(RR4_WL_ENTRY, RR4_LEFT + 1).Value)))
    Dim st As String: st = CellStr(ws.cells(RR4_WL_ENTRY, RR4_LEFT + 2).Value)
    Dim tg As Variant: tg = ws.cells(RR4_WL_ENTRY, RR4_LEFT + 3).Value
    If tk = "" Then Exit Sub
    If Not IsNumeric(tg) Or IsEmpty(tg) Then Exit Sub
    If CDbl(tg) <= 0 Then Exit Sub

    ' first free saved row
    Dim r As Long, dest As Long
    For r = RR4_WL_FIRST To RR4_WL_LAST
        If CellStr(ws.cells(r, RR4_LEFT + 1).Value) = "" Then dest = r: Exit For
    Next r
    If dest = 0 Then
        Call NavNotify("WATCHLIST full (" & (RR4_WL_LAST - RR4_WL_FIRST + 1) & " rows) - double-click a row to delete one", True)
        Exit Sub
    End If
    ws.cells(dest, RR4_LEFT + 1).NumberFormat = "@"
    ws.cells(dest, RR4_LEFT + 1).Value = tk
    ws.cells(dest, RR4_LEFT + 2).NumberFormat = "@"
    ws.cells(dest, RR4_LEFT + 2).Value = st
    ws.cells(dest, RR4_LEFT + 3).NumberFormat = "#,##0.00"
    ws.cells(dest, RR4_LEFT + 3).Value = CDbl(tg)
    Call RefreshWatchlistRow(ws, dest)
    Call WatchlistPaintEntryRow(ws)
    ws.cells(RR4_WL_ENTRY, RR4_LEFT + 1).Select
    Call NavNotify("WATCHLIST + " & tk & " @ " & Format(CDbl(tg), "#,##0.00"))
End Sub

' Called by the sheet code on a double-click inside the saved rows: drop
' that entry and close the gap.
Public Sub WatchlistDeleteRow(ws As Worksheet, ByVal r As Long)
    If r < RR4_WL_FIRST Or r > RR4_WL_LAST Then Exit Sub
    Dim tk As String: tk = CellStr(ws.cells(r, RR4_LEFT + 1).Value)
    If tk = "" Then Exit Sub
    Dim k As Long
    For k = r To RR4_WL_LAST - 1
        ws.cells(k, RR4_LEFT + 1).Value = ws.cells(k + 1, RR4_LEFT + 1).Value
        ws.cells(k, RR4_LEFT + 2).Value = ws.cells(k + 1, RR4_LEFT + 2).Value
        ws.cells(k, RR4_LEFT + 3).Value = ws.cells(k + 1, RR4_LEFT + 3).Value
    Next k
    ws.Range(ws.cells(RR4_WL_LAST, RR4_LEFT + 1), ws.cells(RR4_WL_LAST, RR4_LEFT + 3)).ClearContents
    For k = r To RR4_WL_LAST
        Call RefreshWatchlistRow(ws, k)
    Next k
    Call NavNotify("WATCHLIST - " & tk & " removed")
End Sub

' Live price into E and the lit / unlit state of one saved row.
Public Sub RefreshWatchlistRow(ws As Worksheet, ByVal r As Long)
    If r < RR4_WL_FIRST Or r > RR4_WL_LAST Then Exit Sub
    Dim tk As String: tk = UCase(CellStr(ws.cells(r, RR4_LEFT + 1).Value))
    Dim tgt As Double: tgt = NumOr0(ws.cells(r, RR4_LEFT + 3).Value)
    Dim px As Double
    If tk <> "" Then
        On Error Resume Next
        px = GetStockPrice(tk)
        On Error GoTo 0
    End If
    Dim cells4 As Range
    Set cells4 = ws.Range(ws.cells(r, RR4_LEFT + 1), ws.cells(r, RR4_LEFT + 4))
    cells4.NumberFormat = "General"
    ws.cells(r, RR4_LEFT + 1).NumberFormat = "@"
    ws.cells(r, RR4_LEFT + 2).NumberFormat = "@"
    ws.cells(r, RR4_LEFT + 3).NumberFormat = "#,##0.00"
    ws.cells(r, RR4_LEFT + 4).NumberFormat = "#,##0.00"
    ws.cells(r, RR4_LEFT + 1).HorizontalAlignment = xlCenter
    ws.cells(r, RR4_LEFT + 2).HorizontalAlignment = xlLeft
    ws.cells(r, RR4_LEFT + 3).HorizontalAlignment = xlCenter
    ws.cells(r, RR4_LEFT + 4).HorizontalAlignment = xlCenter
    With ws.cells(r, RR4_LEFT + 4)
        If px > 0 Then .Value = px Else .Value = IIf(tk = "", "", "-")
    End With
    ' lit when the price has come down to the entry target
    Dim hit As Boolean: hit = (px > 0 And tgt > 0 And px <= tgt)
    If hit Then
        cells4.Interior.Color = RGB(60, 30, 0)
        cells4.Font.Color = RR4_ACCENT
        cells4.Font.Bold = True
    Else
        cells4.Interior.Color = RGB(0, 0, 0)
        cells4.Font.Color = RGB(221, 221, 221)
        cells4.Font.Bold = False
        ws.cells(r, RR4_LEFT + 1).Font.Color = RR4_ACCENT
        ws.cells(r, RR4_LEFT + 1).Font.Bold = True
    End If
End Sub

' Slice n of the donut (1-based, cycles after RR4_PALETTE_N).
Private Function DonutColor(ByVal n As Long) As Long
    Select Case ((n - 1) Mod RR4_PALETTE_N) + 1
        Case 1:  DonutColor = RGB(255, 192, 0)      ' amber
        Case 2:  DonutColor = RGB(0, 170, 255)      ' sky blue
        Case 3:  DonutColor = RGB(0, 200, 120)      ' green
        Case 4:  DonutColor = RGB(230, 80, 80)      ' red
        Case 5:  DonutColor = RGB(170, 110, 255)    ' violet
        Case 6:  DonutColor = RGB(255, 130, 40)     ' orange
        Case 7:  DonutColor = RGB(0, 210, 210)      ' teal
        Case 8:  DonutColor = RGB(240, 90, 180)     ' pink
        Case 9:  DonutColor = RGB(190, 210, 60)     ' lime
        Case 10: DonutColor = RGB(90, 130, 255)     ' indigo
        Case 11: DonutColor = RGB(255, 220, 120)    ' pale gold
        Case 12: DonutColor = RGB(160, 160, 160)    ' grey
    End Select
End Function

' ================================================================
'  Read-only accessors for other modules (Attach, TickerInsight) so
'  they never hard-code where the RR4 page keeps a value.
' ================================================================
' Keep the InceptionDate / StartingCapital names on the config cells - they
' pointed at S1/S2 until v4.2 moved the pair to T1/T2.
Private Sub PointConfigNames(ws As Worksheet)
    On Error Resume Next
    ThisWorkbook.names("InceptionDate").Delete
    ThisWorkbook.names("StartingCapital").Delete
    ThisWorkbook.names.Add "InceptionDate", ws.Range(RR4_CFG_INC)
    ThisWorkbook.names.Add "StartingCapital", ws.Range(RR4_CFG_CAP)
    ws.Columns(20).Hidden = True
    On Error GoTo 0
End Sub

Public Function RR4FxRate() As Double
    RR4FxRate = GetExRate(ThisWorkbook.Sheets(SH_PORT))
End Function

Public Function RR4PositionCount() As Long
    Dim lastR As Long: lastR = LastPositionRow(ThisWorkbook.Sheets(SH_PORT))
    If lastR >= RR4_POS_FIRST Then RR4PositionCount = lastR - RR4_POS_FIRST + 1
End Function

' ================================================================
'  Small helpers (v4)
' ================================================================
' SWING RISK is typed by hand, so it has to survive the sheet clear.
' Found by its header text, which works on both the v3 layout (row 4,
' column Q) and v4 (row 25, column P).
Private Function ReadSwingRisk(ws As Worksheet) As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    Dim hdrRow As Long, hdrCol As Long, r As Long, c As Long
    For r = 1 To 60
        For c = 1 To 20
            If UCase(CellStr(ws.cells(r, c).Value)) = "SWING RISK" Then
                hdrRow = r: hdrCol = c
                Exit For
            End If
        Next c
        If hdrRow > 0 Then Exit For
    Next r
    If hdrRow > 0 Then
        Dim lastR As Long: lastR = ws.cells(ws.Rows.count, RR4_LEFT + 1).End(xlUp).row
        For r = hdrRow + 1 To lastR
            Dim tk As String: tk = CellStr(ws.cells(r, RR4_LEFT + 1).Value)
            Dim sv As String: sv = CellStr(ws.cells(r, hdrCol).Value)
            If tk <> "" And sv <> "" Then m(tk) = sv
        Next r
    End If
    Set ReadSwingRisk = m
End Function

' HistoryLog column C (cumulative PnL) on the last row dated before today
Private Function PrevDayCumPnL() As Variant
    PrevDayCumPnL = Empty
    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = ThisWorkbook.Sheets(SH_HIST): On Error GoTo 0
    If wsH Is Nothing Then Exit Function
    Dim r As Long
    For r = wsH.cells(wsH.Rows.count, "A").End(xlUp).row To 2 Step -1
        Dim dv As Variant: dv = wsH.cells(r, "A").Value
        If IsDate(dv) Then
            If Int(CDate(dv)) < Date Then
                Dim cv As Variant: cv = wsH.cells(r, "C").Value
                If Not IsError(cv) And Not IsEmpty(cv) And IsNumeric(cv) Then PrevDayCumPnL = CDbl(cv)
                Exit Function
            End If
        End If
    Next r
End Function

' Realized sheet: sum of PNL(TWD) (col H) for sells dated before d (col I)
Private Function RealizedBefore(ByVal d As Date) As Double
    Dim wsR As Worksheet
    On Error Resume Next: Set wsR = ThisWorkbook.Sheets(SH_REAL): On Error GoTo 0
    If wsR Is Nothing Then Exit Function
    Dim r As Long, lastR As Long
    lastR = wsR.cells(wsR.Rows.count, "A").End(xlUp).row
    For r = 2 To lastR
        Dim dv As Variant: dv = wsR.cells(r, "I").Value
        If IsDate(dv) Then
            If Int(CDate(dv)) < d Then RealizedBefore = RealizedBefore + NumOr0(wsR.cells(r, "H").Value)
        End If
    Next r
End Function

Private Function PosCountOf(posData() As Variant) As Long
    On Error Resume Next
    PosCountOf = UBound(posData, 1)
    If Err.Number <> 0 Then PosCountOf = 0
    Err.Clear
    On Error GoTo 0
End Function

Private Function CellStr(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    CellStr = Trim(CStr(v))
End Function

Private Function NumOr0(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then NumOr0 = CDbl(v)
End Function

Private Function FormatShares(ByVal sh As Double) As String
    If sh = Int(sh) Then
        FormatShares = Format(sh, "#,##0")
    Else
        FormatShares = Format(sh, "#,##0.####")
    End If
End Function

Private Function ShortTicker(ByVal t As String) As String
    t = UCase(Trim(t))
    If Right(t, 4) = ".TWO" Then
        t = Left(t, Len(t) - 4)
    ElseIf Right(t, 3) = ".TW" Then
        t = Left(t, Len(t) - 3)
    End If
    ShortTicker = t
End Function

' Green gain / red loss - the convention of the ticker panel, the weight
' bar and the loss-red holdings rows, so the RR4 page reads one way.
Private Function GainLossColor(ByVal v As Double) As Long
    If v > 0 Then
        GainLossColor = RGB(0, 210, 100)
    ElseIf v < 0 Then
        GainLossColor = RGB(255, 80, 80)
    Else
        GainLossColor = RGB(200, 200, 200)
    End If
End Function
' ================================================================
'  Disclaimer (bottom rows)
' ================================================================
Private Sub DrawDisclaimer(ws As Worksheet, startRow As Long)
    Dim r As Long: r = startRow + 3

    On Error Resume Next
    ws.Range(ws.cells(r, RR4_LEFT + 1), ws.cells(r, RR4_LEFT + 16)).UnMerge
    ws.Range(ws.cells(r, RR4_LEFT + 1), ws.cells(r, RR4_LEFT + 16)).Merge
    On Error GoTo 0
    ws.cells(r, RR4_LEFT + 1).Value = "On the way in Medium-High  Beta Between 2-2.5 , Remember alwaus do the Eliminate underperformers"
    ws.cells(r, RR4_LEFT + 1).Font.Color = RR4_ACCENT
    ws.cells(r, RR4_LEFT + 1).Font.Italic = True
    ws.cells(r, RR4_LEFT + 1).HorizontalAlignment = xlLeft

    r = r + 1
    On Error Resume Next
    ws.Range(ws.cells(r, RR4_LEFT + 1), ws.cells(r, RR4_LEFT + 16)).UnMerge
    ws.Range(ws.cells(r, RR4_LEFT + 1), ws.cells(r, RR4_LEFT + 16)).Merge
    On Error GoTo 0
    ws.cells(r, RR4_LEFT + 1).Value = "FOCUS ON WHAT MY ACTIONS ARE & DO YOUR OWN WORK ! => WHO SAYS YOU CAN'T FIND BETTER TRADE-IDEAS ? => Narrative + Money Flow + Technicals = Valuation Skyrocket"
    ws.cells(r, RR4_LEFT + 1).Font.Color = RR4_ACCENT
    ws.cells(r, RR4_LEFT + 1).Font.Italic = True
    ws.cells(r, RR4_LEFT + 1).HorizontalAlignment = xlLeft
End Sub

Private Sub CalcPositions(positions As Object, exRate As Double, _
                           ByRef posData() As Variant, _
                           ByRef totalMkt As Double, _
                           ByRef totalCost As Double, _
                           ByRef totalUnrl As Double, _
                           ByRef posCount As Long)
    totalMkt = 0: totalCost = 0: totalUnrl = 0: posCount = 0

    ' ------------------------------------------------------------
    Dim tkv As Variant
    For Each tkv In positions.keys
        Dim d0 As Variant: d0 = positions(tkv)
        If d0(0) > 0 Then
            Dim pp0 As Long: pp0 = InStr(CStr(tkv), "|")
            Dim t0 As String: t0 = IIf(pp0 > 0, Left(CStr(tkv), pp0 - 1), CStr(tkv))
            Dim px0 As Double: px0 = GetStockPrice(t0)
            If px0 > 0 Then
                If GetCurrencyType(t0) = "USD" Then
                    totalMkt = totalMkt + px0 * d0(0) * exRate
                Else
                    totalMkt = totalMkt + px0 * d0(0)
                End If
            End If
        End If
    Next tkv

    ' ------------------------------------------------------------
    Dim validCount As Long: validCount = 0
    For Each tkv In positions.keys
        If positions(tkv)(0) > 0 Then validCount = validCount + 1
    Next tkv
    If validCount = 0 Then Exit Sub

    ' ------------------------------------------------------------
    Dim sortedKeys() As String
    ReDim sortedKeys(0 To validCount - 1)
    Dim ki As Long: ki = 0

    Dim brokerOrder(0 To 1) As String
    brokerOrder(0) = "Default"
    brokerOrder(1) = ChrW(&H570B) & ChrW(&H6CF0) & ChrW(&H4E16) & ChrW(&H83EF)  ' Cathay United

    Dim bi As Long
    For bi = 0 To 1
        Dim bName As String: bName = brokerOrder(bi)
        Dim twStart As Long: twStart = ki

        ' -- broker -- x --
        For Each tkv In positions.keys
            If positions(tkv)(0) > 0 Then
                Dim pp1 As Long: pp1 = InStr(CStr(tkv), "|")
                If pp1 > 0 Then
                    Dim t1 As String: t1 = Left(CStr(tkv), pp1 - 1)
                    Dim b1 As String: b1 = Mid(CStr(tkv), pp1 + 1)
                    If b1 = bName And GetCurrencyType(t1) = "TWD" Then
                        sortedKeys(ki) = CStr(tkv): ki = ki + 1
                    End If
                End If
            End If
        Next tkv

        ' ------------------------------------------------------------
        Dim sa As Long, sb As Long, tmp1 As String
        For sa = twStart To ki - 2
            For sb = sa + 1 To ki - 1
                Dim ppA As Long: ppA = InStr(sortedKeys(sa), "|")
                Dim ppB As Long: ppB = InStr(sortedKeys(sb), "|")
                Dim cA As String: cA = Left(sortedKeys(sa), ppA - 1)
                Dim cB As String: cB = Left(sortedKeys(sb), ppB - 1)
                If IsNumeric(cA) And IsNumeric(cB) Then
                    If val(cA) > val(cB) Then
                        tmp1 = sortedKeys(sa): sortedKeys(sa) = sortedKeys(sb): sortedKeys(sb) = tmp1
                    End If
                End If
            Next sb
        Next sa

        Dim usStart As Long: usStart = ki

        ' -- broker --
        For Each tkv In positions.keys
            If positions(tkv)(0) > 0 Then
                Dim pp2 As Long: pp2 = InStr(CStr(tkv), "|")
                If pp2 > 0 Then
                    Dim t2 As String: t2 = Left(CStr(tkv), pp2 - 1)
                    Dim b2 As String: b2 = Mid(CStr(tkv), pp2 + 1)
                    If b2 = bName And GetCurrencyType(t2) = "USD" Then
                        sortedKeys(ki) = CStr(tkv): ki = ki + 1
                    End If
                End If
            End If
        Next tkv

        ' ------------------------------------------------------------
        Dim tmp2 As String
        For sa = usStart To ki - 2
            For sb = sa + 1 To ki - 1
                Dim ppA2 As Long: ppA2 = InStr(sortedKeys(sa), "|")
                Dim ppB2 As Long: ppB2 = InStr(sortedKeys(sb), "|")
                Dim cA2 As String: cA2 = Left(sortedKeys(sa), ppA2 - 1)
                Dim cB2 As String: cB2 = Left(sortedKeys(sb), ppB2 - 1)
                If cA2 > cB2 Then
                    tmp2 = sortedKeys(sa): sortedKeys(sa) = sortedKeys(sb): sortedKeys(sb) = tmp2
                End If
            Next sb
        Next sa
    Next bi

    ' ------------------------------------------------------------
    Dim addedKeys As Object: Set addedKeys = CreateObject("Scripting.Dictionary")
    Dim si2 As Long
    For si2 = 0 To ki - 1: addedKeys(sortedKeys(si2)) = 1: Next si2
    For Each tkv In positions.keys
        If positions(tkv)(0) > 0 Then
            If Not addedKeys.Exists(CStr(tkv)) Then
                sortedKeys(ki) = CStr(tkv): ki = ki + 1
            End If
        End If
    Next tkv

    ' ------------------------------------------------------------
    ReDim posData(1 To validCount, 1 To 16)
    posCount = 0

    Dim si As Long
    For si = 0 To validCount - 1
        Dim compositeKey As String: compositeKey = sortedKeys(si)
        Dim pipePos As Long: pipePos = InStr(compositeKey, "|")
        Dim tickerStr As String: tickerStr = Left(compositeKey, pipePos - 1)
        Dim brokerPart As String: brokerPart = Mid(compositeKey, pipePos + 1)

        Dim d As Variant: d = positions(compositeKey)
        Dim netQty As Double: netQty = d(0)
        If netQty <= 0 Then GoTo NextSi

        posCount = posCount + 1
        Dim p As Long: p = posCount

        Dim totCost As Double: totCost = d(2)
        Dim entryDt As Date
        If d(5) > 0 Then entryDt = CDate(d(5)) Else entryDt = Date

        Dim AvgCost As Double
        If netQty > 0 Then AvgCost = totCost / netQty

        Dim curType   As String: curType = GetCurrencyType(tickerStr)
        Dim livePrice As Double: livePrice = GetStockPrice(tickerStr)

        Dim mktVal As Double, unrlPnl As Double
        If livePrice > 0 Then
            mktVal = livePrice * netQty
            unrlPnl = mktVal - totCost
        Else
            mktVal = totCost: unrlPnl = 0
        End If

        Dim mktTWD As Double, unrlTWD As Double
        If curType = "USD" Then
            mktTWD = mktVal * exRate: unrlTWD = unrlPnl * exRate
        Else
            mktTWD = mktVal: unrlTWD = unrlPnl
        End If

        totalCost = totalCost + IIf(curType = "USD", totCost * exRate, totCost)
        totalUnrl = totalUnrl + unrlTWD

        Dim chgPct As Double
        If AvgCost > 0 And livePrice > 0 Then chgPct = (livePrice - AvgCost) / AvgCost

        posData(p, 1) = tickerStr
        posData(p, 2) = tickerStr
        posData(p, 3) = GetCompanyName(tickerStr)
        posData(p, 4) = entryDt
        posData(p, 5) = Date - entryDt + 1
        posData(p, 6) = brokerPart
        posData(p, 7) = CStr(d(6))
        posData(p, 8) = mktTWD
        posData(p, 9) = netQty
        posData(p, 10) = AvgCost
        posData(p, 11) = livePrice
        posData(p, 12) = chgPct
        posData(p, 13) = unrlTWD
        posData(p, 14) = d(9)
        posData(p, 15) = d(9)
        posData(p, 16) = d(7)
NextSi:
    Next si
End Sub
' ================================================================
'  Position builder - FIFO cost basis (2026-08-19)
' ----------------------------------------------------------------
'  Was weighted-average cost: SELL reduced the running total d(2) by
'  avg*shares where avg = d(2)/d(0). That let this dashboard's ENTRY PX /
'  UNRL PNL diverge from the Realized sheet (Attach.CalculateRealizedPnL)
'  and Ticker Insight (TickerInsight.BuildFIFOHistory) - both already walk
'  the same Transactions data FIFO, so the same ticker with two BUY lots
'  at different prices could show three different cost bases across the
'  workbook. This now keeps an explicit FIFO lot queue per Ticker|Broker
'  key (oldest lot consumed first on SELL), with ADJUSTCOST allocated
'  proportionally across open lots by share count - mirroring
'  Attach.CalculateRealizedPnL's own lot handling so the two agree.
'
'  Return shape is unchanged: still Ticker|Broker -> the same 11-element
'  Array CalcPositions already reads by position. d(0)/d(2) are still the
'  aggregate remaining shares/cost (now the FIFO sum, not a weighted
'  running total); d(5) is now the OLDEST still-open lot's date instead
'  of "first BUY ever, reset to 0 on full close" - same value while a
'  position stays open continuously, but correct again after a position
'  is closed and reopened with a later lot. d(11) is new: the raw FIFO
'  lot Collection (date, shares, cost/share), exposed for any future
'  caller that needs true per-lot detail instead of the aggregate.
' ================================================================
Private Function BuildPositions(wsTr As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = wsTr.cells(wsTr.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Set BuildPositions = dict: Exit Function

    Dim data As Variant
    data = wsTr.Range("B2:N" & lastRow).Value   ' B..N

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim Ticker As String: Ticker = CStr(data(i, 2))
        Dim Action As String: Action = CStr(data(i, 3))
        Dim shares As Double: shares = val(data(i, 4))
        Dim netAmt As Double: netAmt = val(data(i, 8))
        Dim Sector As String: Sector = CStr(data(i, 9))
        Dim pTgt   As Double: pTgt = val(data(i, 10))
        Dim strat  As String: strat = CStr(data(i, 11))
        Dim Beta   As Double: Beta = val(data(i, 12))
        Dim broker As String
        On Error Resume Next
        broker = Trim(CStr(data(i, 13)))
        On Error GoTo 0
        If broker = "" Then broker = "Default"

        If Ticker <> "" Then
            Dim posKey As String: posKey = Ticker & "|" & broker
            If Not dict.Exists(posKey) Then
                dict.Add posKey, Array(0#, 0#, 0#, 0#, 0#, 0#, "", 0#, "", 0#, broker, New Collection)
            End If
            Dim d As Variant: d = dict(posKey)
            If pTgt > 0 Then d(7) = pTgt
            If strat <> "" Then d(8) = strat
            If Beta <> 0 Then d(9) = Beta
            If Sector <> "" Then d(6) = Sector
            d(10) = broker

            Dim lots As Collection: Set lots = d(11)
            Dim tDate As Variant: tDate = data(i, 1)

            Select Case UCase(Action)
                Case "BUY"
                    Dim cps As Double: If shares > 0 Then cps = netAmt / shares Else cps = 0
                    Dim lotDateSer As Long: lotDateSer = IIf(IsDate(tDate), CLng(CDate(tDate)), 0)
                    lots.Add Array(lotDateSer, shares, cps)

                Case "SELL"
                    Dim remaining As Double: remaining = shares
                    Do While remaining > 0.0001 And lots.count > 0
                        Dim frontLot As Variant: frontLot = lots(1)
                        Dim lotSh As Double: lotSh = frontLot(1)
                        Dim lotCps As Double: lotCps = frontLot(2)
                        If lotSh <= remaining + 0.0001 Then
                            remaining = remaining - lotSh
                            lots.Remove 1
                        Else
                            lots.Add Array(frontLot(0), lotSh - remaining, lotCps), Before:=1
                            lots.Remove 2
                            remaining = 0
                        End If
                    Loop

                Case "ADJUSTCOST"
                    ' Spread proportionally across open lots by share count,
                    ' matching Attach.CalculateRealizedPnL's ADJUSTCOST rule.
                    Dim totSh As Double: totSh = 0
                    Dim li As Long
                    For li = 1 To lots.count: totSh = totSh + lots(li)(1): Next li
                    If totSh > 0 Then
                        Dim adjPerSh As Double: adjPerSh = netAmt / totSh
                        For li = 1 To lots.count
                            Dim lotArr As Variant: lotArr = lots(li)
                            lotArr(2) = lotArr(2) + adjPerSh
                            If li < lots.count Then
                                lots.Add lotArr, Before:=li + 1
                                lots.Remove li
                            Else
                                lots.Add lotArr
                                lots.Remove li
                            End If
                        Next li
                    End If
            End Select

            ' Recompute the aggregate shares / cost / oldest-lot date from
            ' the queue so d(0)/d(2)/d(5) - the fields CalcPositions already
            ' reads - stay in sync with whatever the FIFO queue now holds.
            Dim aggShares As Double, aggCost As Double, oldestDate As Long
            aggShares = 0: aggCost = 0: oldestDate = 0
            Dim lj As Long
            For lj = 1 To lots.count
                aggShares = aggShares + lots(lj)(1)
                aggCost = aggCost + lots(lj)(1) * lots(lj)(2)
                Dim ld As Long: ld = lots(lj)(0)
                If ld > 0 And (oldestDate = 0 Or ld < oldestDate) Then oldestDate = ld
            Next lj
            d(0) = aggShares
            d(2) = aggCost
            d(5) = oldestDate
            ' d(11) holds an object (the FIFO lot Collection); an object
            ' assignment needs Set - a plain Let-assign raises run-time
            ' error 450 (invalid property assignment) and aborts the whole
            ' dashboard rebuild before the page is drawn.
            Set d(11) = lots

            dict(posKey) = d
        End If
    Next i
    Set BuildPositions = dict
End Function
' ================================================================
'  Portfolio Beta
' ================================================================
Private Function CalcPortBeta(posData() As Variant) As Double
    CalcPortBeta = 0
    If Not IsArray(posData) Then Exit Function
    On Error Resume Next
    Dim n As Long: n = UBound(posData, 1)
    If Err.Number <> 0 Or n < 1 Then Exit Function
    On Error GoTo 0
    
    Dim totalMkt As Double, totalBeta As Double
    Dim i As Long
    For i = 1 To n
        totalMkt = totalMkt + posData(i, 8)
        totalBeta = totalBeta + posData(i, 8) * posData(i, 14)
    Next i
    If totalMkt > 0 Then CalcPortBeta = totalBeta / totalMkt
End Function



' ------------------------------------------------------------
Private Function FindNearestPrice(ByRef dict As Object, ByVal targetDate As Date) As Double
    Dim checkDate As Date
    checkDate = targetDate
    Dim daysBack As Integer
    
    ' ------------------------------------------------------------
    For daysBack = 1 To 10
        checkDate = checkDate - 1
        If dict.Exists(checkDate) Then
            FindNearestPrice = dict(checkDate)
            Exit Function
        End If
    Next daysBack
    
    FindNearestPrice = 0
End Function
' ================================================================
'  Helpers
' ================================================================
Private Function PnLColor(v As Double) As Long
    If v > 0 Then
        PnLColor = RGB(255, 80, 80)
    ElseIf v < 0 Then
        PnLColor = RGB(0, 210, 100)
    Else
        PnLColor = RGB(200, 200, 200)
    End If
End Function

' Muted variant used ONLY by the RR4 holdings rows (columns K/L).
' Profit and flat share one grey; losses get a soft red so the eye lands
' on what is bleeding. Feed it the UNRL PNL value for BOTH columns so a
' short position cannot paint % CHG and PNL in opposite colours.
Private Function PnLColorMuted(v As Double) As Long
    If v < 0 Then
        PnLColorMuted = RGB(255, 130, 130)
    Else
        PnLColorMuted = RGB(200, 200, 200)
    End If
End Function

' USD/TWD from the page's input cell. B2 is where v4 kept it - read once
' as a fallback so the first v4.1 rebuild carries a typed rate across.
Private Function GetExRate(ws As Worksheet) As Double
    ' (the v4 "read B2 as well" fallback is gone: B2 is a nav-bar cell now)
    Dim v As Double
    v = NumOr0(ws.Range(RR4_FX_CELL).Value)
    If v < 20 Or v > 50 Then v = 31.6
    GetExRate = v
End Function

Private Function GetInceptionDate(ws As Worksheet) As Date
    On Error Resume Next
    GetInceptionDate = CDate(ws.Range("InceptionDate").Value)
    On Error GoTo 0
    If GetInceptionDate <= DateSerial(2000, 1, 1) Then GetInceptionDate = DateSerial(2026, 8, 1)
End Function

Private Function GetStartingCapital(ws As Worksheet) As Double
    On Error Resume Next
    GetStartingCapital = val(ws.Range("StartingCapital").Value)
    On Error GoTo 0
    If GetStartingCapital <= 0 Then GetStartingCapital = 600000
End Function

' ================================================================
'  One-time config setup
' ================================================================
Sub SetupPortfolioConfig()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PORT)

    Dim inDate As String
    inDate = InputBox("Inception Date (yyyy/m/d):", "Config", "2026/8/1")
    If Not IsDate(inDate) Then MsgBox "Invalid date format.": Exit Sub

    Dim startCap As String
    startCap = InputBox("Starting Capital (TWD, e.g. 600000):", "Config", "600000")
    If Not IsNumeric(startCap) Then MsgBox "Invalid number.": Exit Sub

    On Error Resume Next
    ThisWorkbook.names("InceptionDate").Delete
    ThisWorkbook.names("StartingCapital").Delete
    On Error GoTo 0

    ws.Range(RR4_CFG_INC).Value = CDate(inDate)
    ws.Range(RR4_CFG_CAP).Value = CDbl(startCap)
    ThisWorkbook.names.Add "InceptionDate", ws.Range(RR4_CFG_INC)
    ThisWorkbook.names.Add "StartingCapital", ws.Range(RR4_CFG_CAP)
    ws.Columns("S").Hidden = True

    MsgBox "Config saved!" & vbCr & "Inception: " & inDate & vbCr & "Capital: " & startCap, vbInformation
End Sub
Private Sub WriteKV(ws As Worksheet, r As Long, label As String, _
                    val As Double, fmt As String, colorPnL As Boolean)
    ws.cells(r, 1).Value = label
    ws.cells(r, 1).Font.Color = RGB(150, 150, 150)
    ws.cells(r, 2).Value = val
    ws.cells(r, 2).NumberFormat = fmt
    ws.cells(r, 2).Font.Bold = True
    If colorPnL Then
        ws.cells(r, 2).Font.Color = PnLColor(val)
    Else
        ws.cells(r, 2).Font.Color = RGB(221, 221, 221)
    End If
End Sub


Sub BuildHoldingsCorrelation()
    Application.ScreenUpdating = False

    Dim wsTr As Worksheet
    Set wsTr = ThisWorkbook.Sheets("Transactions")

    Dim positions As Object
    Set positions = BuildPositions(wsTr)

    ' ------------------------------------------------------------
    Dim uniqTickers As Object: Set uniqTickers = CreateObject("Scripting.Dictionary")
    Dim tkv As Variant
    For Each tkv In positions.keys
        If positions(tkv)(0) > 0 Then
            Dim pkStr As String: pkStr = CStr(tkv)
            Dim barPos As Long: barPos = InStr(pkStr, "|")
            Dim pureT As String
            If barPos > 0 Then pureT = Left(pkStr, barPos - 1) Else pureT = pkStr
            pureT = UCase(Trim(pureT))
            If pureT <> "" And Not uniqTickers.Exists(pureT) Then
                uniqTickers.Add pureT, 1
            End If
        End If
    Next tkv

    Dim tickerCount As Long: tickerCount = uniqTickers.count
    If tickerCount < 2 Then
        Call NavNotify("C!: need at least 2 holdings", True)
        Application.ScreenUpdating = True: Exit Sub
    End If

    ' ------------------------------------------------------------
    Dim rawTickers() As String
    ReDim rawTickers(0 To tickerCount - 1)
    Dim ki As Long: ki = 0
    For Each tkv In uniqTickers.keys
        rawTickers(ki) = CStr(tkv): ki = ki + 1
    Next tkv

    ' ------------------------------------------------------------
    ' ------------------------------------------------------------
    Dim tickers() As String
    ReDim tickers(0 To tickerCount - 1)
    Dim sortIdx() As Long
    ReDim sortIdx(0 To tickerCount - 1)

    ' Step 1: -- x -- index -- J
    Dim twCount As Long: twCount = 0
    Dim i As Long
    For i = 0 To tickerCount - 1
        If GetCurrencyType(rawTickers(i)) = "TWD" Then
            sortIdx(twCount) = i
            twCount = twCount + 1
        End If
    Next i

    ' Step 2: fill US tickers next
    Dim usIdx As Long: usIdx = twCount
    For i = 0 To tickerCount - 1
        If GetCurrencyType(rawTickers(i)) = "USD" Then
            sortIdx(usIdx) = i
            usIdx = usIdx + 1
        End If
    Next i

    ' Step 3: sort TW tickers numerically
    Dim a As Long, b As Long, tmpL As Long
    For a = 0 To twCount - 2
        For b = a + 1 To twCount - 1
            Dim codeA As String: codeA = Replace(Replace(UCase(rawTickers(sortIdx(a))), ".TWO", ""), ".TW", "")
            Dim codeB As String: codeB = Replace(Replace(UCase(rawTickers(sortIdx(b))), ".TWO", ""), ".TW", "")
            If IsNumeric(codeA) And IsNumeric(codeB) Then
                If val(codeA) > val(codeB) Then
                    tmpL = sortIdx(a): sortIdx(a) = sortIdx(b): sortIdx(b) = tmpL
                End If
            End If
        Next b
    Next a

    ' Step 4: sort US tickers alphabetically
    For a = twCount To tickerCount - 2
        For b = a + 1 To tickerCount - 1
            If rawTickers(sortIdx(a)) > rawTickers(sortIdx(b)) Then
                tmpL = sortIdx(a): sortIdx(a) = sortIdx(b): sortIdx(b) = tmpL
            End If
        Next b
    Next a

    ' Step 5: apply sorted order
    For i = 0 To tickerCount - 1
        tickers(i) = rawTickers(sortIdx(i))
    Next i

    ' ------------------------------------------------------------
    Dim priceData() As Object
    ReDim priceData(0 To tickerCount - 1)
    For i = 0 To tickerCount - 1
        Application.StatusBar = "Fetching " & tickers(i) & " (" & (i + 1) & "/" & tickerCount & ")..."
        Set priceData(i) = GetHistoryPrices(tickers(i))
    Next i

    ' ------------------------------------------------------------
    Dim commonDates() As Date
    Dim usedCount As Long
    Call CorrGetCommonDates(priceData, tickerCount, commonDates, usedCount)

    If usedCount < 4 Then
        Dim diagMsg As String
        diagMsg = "Insufficient common trading days (" & usedCount & " day(s))." & vbCr & vbCr & _
                 "History row count per ticker:" & vbCr
        Dim dIdx As Long
        For dIdx = 0 To tickerCount - 1
            Dim dcnt As Long: dcnt = 0
            If Not priceData(dIdx) Is Nothing Then dcnt = priceData(dIdx).count
            diagMsg = diagMsg & "  " & tickers(dIdx) & " : " & dcnt & " rows"
            If dcnt = 0 Then diagMsg = diagMsg & "  <-- API returned nothing"
            diagMsg = diagMsg & vbCr
        Next dIdx
        diagMsg = diagMsg & vbCr & "Tip: re-import Attach module, save / close / reopen workbook, retry."
        Debug.Print diagMsg
        Call NavNotify("C!: only " & usedCount & " common trading days - per-ticker counts in the Immediate window", True)
        Application.ScreenUpdating = True: Exit Sub
    End If

    ' ------------------------------------------------------------
        ' ------------------------------------------------------------
    Dim retCount As Long: retCount = usedCount - 1
    Dim returns() As Double
    ReDim returns(0 To tickerCount - 1, 0 To retCount - 1)

    Dim j As Long
    For i = 0 To tickerCount - 1
        For j = 0 To retCount - 1
            Dim px1 As Double: px1 = 0
            Dim px2 As Double: px2 = 0
            On Error Resume Next
            px1 = priceData(i)(commonDates(j))
            px2 = priceData(i)(commonDates(j + 1))
            On Error GoTo 0
            If px1 > 0 And px2 > 0 Then returns(i, j) = Log(px2 / px1)
        Next j
    Next i

    ' ------------------------------------------------------------
    Dim corrMatrix() As Double
    ReDim corrMatrix(0 To tickerCount - 1, 0 To tickerCount - 1)
    Dim r As Long, c As Long
    For r = 0 To tickerCount - 1
        For c = 0 To tickerCount - 1
            corrMatrix(r, c) = CorrPearson(returns, r, c, retCount)
        Next c
    Next r

    ' ------------------------------------------------------------
    Call HoldingsCorrRenderSheet(tickers, tickerCount, corrMatrix, usedCount, commonDates)

    Application.ScreenUpdating = True
    Application.StatusBar = "HoldingsCorr updated: " & Format(Now, "hh:mm:ss")
    Call NavNotify("C! done " & Format(Now, "hh:mm:ss") & " - holdings correlation updated")
End Sub

' ------------------------------------------------------------
Private Sub HoldingsCorrRenderSheet(tickers() As String, tickerCount As Long, _
                                     corrMatrix() As Double, usedCount As Long, _
                                     commonDates() As Date)
    Dim wsC As Worksheet
    On Error Resume Next: Set wsC = ThisWorkbook.Sheets("HoldingsCorr"): On Error GoTo 0
    If wsC Is Nothing Then
        Set wsC = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsC.Name = "HoldingsCorr"
    End If

    Call NavStrip(wsC)
    wsC.cells.Clear
    With wsC.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
    wsC.Activate
    ActiveWindow.DisplayGridlines = False

    ' ------------------------------------------------------------
    Dim retCount As Long: retCount = usedCount - 1
    With wsC.cells(1, 1)
        .Value = "HOLDINGS CORRELATION  |  " & retCount & "-DAY  |  " & _
                 Format(commonDates(0), "m/d/yy") & " ~ " & _
                 Format(commonDates(usedCount - 1), "m/d/yy") & _
                 "  |  " & Format(Now, "hh:mm:ss")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 11
    End With
    With wsC.Range(wsC.cells(2, 1), wsC.cells(2, tickerCount + 2))
        .Interior.Color = RGB(255, 192, 0)
        .RowHeight = 3
    End With

    ' ------------------------------------------------------------
    Dim HDR As Long: HDR = 4
    wsC.Columns(1).ColumnWidth = 12
    Dim i As Long
    For i = 0 To tickerCount - 1
        wsC.Columns(i + 2).ColumnWidth = 7
    Next i

    ' ------------------------------------------------------------
    With wsC.cells(HDR, 1)
         .Value = "Row / Col"
        .Font.Color = RGB(100, 100, 100)
        .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter
    End With

    For i = 0 To tickerCount - 1
        With wsC.cells(HDR, i + 2)
            .Value = tickers(i)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 8
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
        wsC.Rows(HDR).RowHeight = 20
    Next i

    ' ------------------------------------------------------------
    Dim j As Long
    For i = 0 To tickerCount - 1
        Dim rn As Long: rn = HDR + 1 + i
        wsC.Rows(rn).RowHeight = 18

        ' ------------------------------------------------------------
        With wsC.cells(rn, 1)
            .Value = tickers(i)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(10, 10, 10)
        End With

        For j = 0 To tickerCount - 1
            Dim cv As Double: cv = corrMatrix(i, j)
            Dim bgC As Long, fgC As Long

            ' ------------------------------------------------------------
            If i = j Then
                ' ------------------------------------------------------------
                bgC = RGB(40, 28, 0)
                fgC = RGB(255, 192, 0)
            ElseIf cv >= 0.2 Then
                ' ------------------------------------------------------------
                If cv >= 0.7 Then
                    bgC = RGB(100, 0, 0):   fgC = RGB(255, 130, 130)
                ElseIf cv >= 0.4 Then
                    bgC = RGB(70, 0, 0):    fgC = RGB(230, 100, 100)
                Else
                    bgC = RGB(45, 5, 5):    fgC = RGB(210, 80, 80)
                End If
            ElseIf cv <= -0.2 Then
                ' ------------------------------------------------------------
                If cv <= -0.7 Then
                    bgC = RGB(0, 70, 20):   fgC = RGB(80, 255, 140)
                ElseIf cv <= -0.4 Then
                    bgC = RGB(0, 50, 15):   fgC = RGB(60, 220, 110)
                Else
                    bgC = RGB(0, 35, 10):   fgC = RGB(40, 200, 90)
                End If
            Else
                ' ------------------------------------------------------------
                bgC = RGB(18, 18, 18):      fgC = RGB(150, 150, 150)
            End If

            With wsC.cells(rn, j + 2)
                .Value = Round(cv, 2)
                .NumberFormat = "0.00"
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 8
                .Interior.Color = bgC
                .Font.Color = fgC
                ' ------------------------------------------------------------
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = RGB(0, 0, 0)
                    .Weight = xlThin
                End With
            End With
        Next j
    Next i

    ' ------------------------------------------------------------
    Dim LR As Long: LR = HDR + tickerCount + 2
    wsC.cells(LR, 1).Value = "LEGEND"
    wsC.cells(LR, 1).Font.Color = RGB(255, 192, 0)
    wsC.cells(LR, 1).Font.Bold = True

    Dim leg As Variant
    leg = Array( _
        Array(">= 0.70  High +corr", RGB(100, 0, 0), RGB(255, 130, 130)), _
        Array("0.40~0.69 Mid +corr", RGB(70, 0, 0), RGB(230, 100, 100)), _
        Array("0.20~0.39 Low +corr", RGB(45, 5, 5), RGB(210, 80, 80)), _
        Array("-0.19~0.19 Neutral", RGB(18, 18, 18), RGB(150, 150, 150)), _
        Array("-0.20~-0.39 Low -corr", RGB(0, 35, 10), RGB(40, 200, 90)), _
        Array("-0.40~-0.69 Mid -corr", RGB(0, 50, 15), RGB(60, 220, 110)), _
        Array("<= -0.70 High -corr", RGB(0, 70, 20), RGB(80, 255, 140)))

    Dim li As Long
    For li = 0 To UBound(leg)
        With wsC.cells(LR + 1 + li, 1)
            .Value = leg(li)(0)
            .Interior.Color = leg(li)(1)
            .Font.Color = leg(li)(2)
            .Font.Bold = True
            .Font.Size = 9
        End With
        wsC.Rows(LR + 1 + li).RowHeight = 16
    Next li

    ' ------------------------------------------------------------

    Call NavAdd(wsC, "C")
    wsC.Activate
End Sub
' ------------------------------------------------------------
Private Sub CorrGetCommonDates(priceData() As Object, tickerCount As Long, _
                                ByRef commonDates() As Date, ByRef usedCount As Long)
    Dim allDates As Object: Set allDates = CreateObject("Scripting.Dictionary")
    Dim dkey As Variant

    ' ------------------------------------------------------------
    For Each dkey In priceData(0).keys
        allDates(dkey) = 1
    Next dkey

    ' ------------------------------------------------------------
    Dim i As Long
    For i = 1 To tickerCount - 1
        Dim toRemove() As Variant
        Dim removeCount As Long: removeCount = 0
        ReDim toRemove(0 To allDates.count)
        For Each dkey In allDates.keys
            If Not priceData(i).Exists(dkey) Then
                toRemove(removeCount) = dkey
                removeCount = removeCount + 1
            End If
        Next dkey
        Dim k As Long
        For k = 0 To removeCount - 1
            allDates.Remove toRemove(k)
        Next k
    Next i

    Dim dateCount As Long: dateCount = allDates.count
    If dateCount = 0 Then usedCount = 0: Exit Sub

    ' ------------------------------------------------------------
    Dim sortedDates() As Date
    ReDim sortedDates(0 To dateCount - 1)
    Dim di As Long: di = 0
    For Each dkey In allDates.keys
        sortedDates(di) = CDate(dkey): di = di + 1
    Next dkey

    Dim a As Long, b As Long, tmpD As Date
    For a = 0 To dateCount - 2
        For b = a + 1 To dateCount - 1
            If sortedDates(a) > sortedDates(b) Then
                tmpD = sortedDates(a): sortedDates(a) = sortedDates(b): sortedDates(b) = tmpD
            End If
        Next b
    Next a

    ' ------------------------------------------------------------
    ' 180 trading-day window: 181 price points yield 180 log returns.
    ' Short-history tickers fall back to whatever common days exist.
    Dim wantPx As Long: wantPx = 181
    Dim startIdx As Long: startIdx = IIf(dateCount > wantPx, dateCount - wantPx, 0)
    usedCount = dateCount - startIdx

    ReDim commonDates(0 To usedCount - 1)
    For i = 0 To usedCount - 1
        commonDates(i) = sortedDates(startIdx + i)
    Next i
End Sub

' ------------------------------------------------------------
Private Function CorrPearson(returns() As Double, r1 As Long, r2 As Long, n As Long) As Double
    If n < 2 Then CorrPearson = 0: Exit Function

    Dim sumX As Double, sumY As Double, sumXY As Double
    Dim sumX2 As Double, sumY2 As Double, j As Long

    For j = 0 To n - 1
        Dim x As Double: x = returns(r1, j)
        Dim y As Double: y = returns(r2, j)
        sumX = sumX + x:   sumY = sumY + y
        sumXY = sumXY + x * y
        sumX2 = sumX2 + x * x
        sumY2 = sumY2 + y * y
    Next j

    Dim denom As Double
    denom = Sqr((n * sumX2 - sumX ^ 2) * (n * sumY2 - sumY ^ 2))
    If denom = 0 Then CorrPearson = 0 Else CorrPearson = (n * sumXY - sumX * sumY) / denom
End Function

' ------------------------------------------------------------
Private Sub CorrRenderSheet(tickers() As String, tickerCount As Long, _
                             corrMatrix() As Double, usedCount As Long, _
                             commonDates() As Date)
    Dim wsC As Worksheet
    On Error Resume Next: Set wsC = ThisWorkbook.Sheets("HoldingsCorr"): On Error GoTo 0
    If wsC Is Nothing Then
        Set wsC = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsC.Name = "HoldingsCorr"
    End If

    wsC.cells.Clear
    With wsC.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 10
    End With
    wsC.Activate
    ActiveWindow.DisplayGridlines = False

    ' ------------------------------------------------------------
    Dim retCount As Long: retCount = usedCount - 1
    With wsC.cells(1, 1)
        .Value = "CORRELATION MATRIX  |  " & retCount & "-DAY RETURNS  |  " & _
                 Format(commonDates(0), "m/d/yy") & " ~ " & _
                 Format(commonDates(usedCount - 1), "m/d/yy") & _
                 "  |  Updated: " & Format(Now, "hh:mm:ss")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 12
    End With
    With wsC.Range(wsC.cells(2, 1), wsC.cells(2, tickerCount + 3))
        .Interior.Color = RGB(255, 192, 0)
        .RowHeight = 3
    End With

    ' ------------------------------------------------------------
    Dim HDR As Long: HDR = 4
    With wsC.cells(HDR, 1)
         .Value = "ROW / COL"
        .Font.Color = RGB(150, 150, 150)
        .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter
    End With

    Dim i As Long, j As Long
    For i = 0 To tickerCount - 1
        With wsC.cells(HDR, i + 2)
            .Value = tickers(i)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next i

    ' ------------------------------------------------------------
    For i = 0 To tickerCount - 1
        Dim rn As Long: rn = HDR + 1 + i
        wsC.Rows(rn).RowHeight = 20

        With wsC.cells(rn, 1)
            .Value = tickers(i)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(10, 10, 10)
        End With

        For j = 0 To tickerCount - 1
            Dim cv As Double: cv = corrMatrix(i, j)
            Dim bgC As Long, fgC As Long

            If i = j Then                                  ' -- 﨤 -- u -- G -- ۤv -- P -- ۤv = 1.00
                bgC = RGB(40, 30, 0):   fgC = RGB(255, 192, 0)
            ElseIf cv >= 0.7 Then                          ' -- ץ --
                bgC = RGB(90, 10, 10):  fgC = RGB(255, 100, 100)
            ElseIf cv >= 0.4 Then                          ' -- ץ --
                bgC = RGB(55, 20, 0):   fgC = RGB(230, 140, 70)
            ElseIf cv >= 0.1 Then                          ' -- C -- ץ --
                bgC = RGB(22, 22, 22):  fgC = RGB(200, 200, 160)
            ElseIf cv >= -0.1 Then                         ' ------------------------------------------------------------
                bgC = RGB(15, 15, 15):  fgC = RGB(170, 170, 170)
            ElseIf cv >= -0.4 Then                         ' -- C -- ׭t --
                bgC = RGB(0, 22, 28):   fgC = RGB(80, 200, 200)
            Else                                           ' -- ׭t --
                bgC = RGB(0, 45, 20):   fgC = RGB(0, 210, 100)
            End If

            With wsC.cells(rn, j + 2)
                .Value = Round(cv, 2)
                .NumberFormat = "0.00"
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = bgC
                .Font.Color = fgC
            End With
        Next j
    Next i

    ' ------------------------------------------------------------
    Dim LR As Long: LR = HDR + tickerCount + 3
    wsC.cells(LR, 1).Value = "COLOR LEGEND"
    wsC.cells(LR, 1).Font.Color = RGB(255, 192, 0)
    wsC.cells(LR, 1).Font.Bold = True

    Dim leg As Variant
    leg = Array( _
        Array("? 0.70   HIGH POSITIVE", RGB(90, 10, 10), RGB(255, 100, 100)), _
        Array("0.40 ~ 0.69   MODERATE POSITIVE", RGB(55, 20, 0), RGB(230, 140, 70)), _
        Array("0.10 ~ 0.39   LOW POSITIVE", RGB(22, 22, 22), RGB(200, 200, 160)), _
        Array("-0.09 ~ 0.09   NEUTRAL", RGB(15, 15, 15), RGB(170, 170, 170)), _
        Array("-0.10 ~ -0.39   LOW NEGATIVE", RGB(0, 22, 28), RGB(80, 200, 200)), _
        Array("? -0.40   HIGH NEGATIVE", RGB(0, 45, 20), RGB(0, 210, 100)) _
    )

    Dim li As Long
    For li = 0 To UBound(leg)
        With wsC.cells(LR + 1 + li, 1)
            .Value = leg(li)(0)
            .Interior.Color = leg(li)(1)
            .Font.Color = leg(li)(2)
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
        End With
        wsC.Rows(LR + 1 + li).RowHeight = 18
    Next li

    ' ------------------------------------------------------------
    wsC.Columns(1).ColumnWidth = 20
    Dim col As Long
    For col = 2 To tickerCount + 1
        wsC.Columns(col).ColumnWidth = IIf(tickerCount > 10, 8, 10)
    Next col

    wsC.Activate
End Sub

' ================================================================
'  180D VOLATILITY - per-stock std dev + market-value-weighted
'  portfolio std dev
' ----------------------------------------------------------------
'  Nav command V! (was the VOLATILITY button). Not run automatically
'  from RebuildPortfolioDashboard - each holding needs its own ~1y
'  Yahoo history call, too slow to fire on every dashboard refresh.
'
'  Per-stock std dev: last 180 trading days of SIMPLE daily returns,
'  each ticker on its OWN most-recent window (independent of the other
'  holdings). Annualized by * SQR(252). Tickers with fewer than 180
'  days of history just use whatever is available - the days-used
'  column shows the real count instead of erroring out.
'
'  Portfolio std dev: reuses CorrGetCommonDates - the same date-
'  alignment routine BuildHoldingsCorrelation (HOLDINGS CORR button)
'  already uses - to find the trading days common to ALL current
'  holdings (the shortest-history holding caps the window, same as
'  HOLDINGS CORR). Each holding's SIMPLE daily return on those shared
'  dates is weighted by its CURRENT market value (TWD) and summed per
'  day; the std dev of that weighted-sum series, annualized, equals
'  SQR(w' * Cov * w) without ever building or storing an explicit
'  covariance matrix. SIMPLE returns are used here (not the log
'  returns HOLDINGS CORR uses for its own, unrelated correlation
'  matrix) because a portfolio's simple return really does equal the
'  weighted sum of its holdings' simple returns - that identity does
'  not hold for log returns.
' ================================================================
Sub UpdatePortfolioVolatility()
    Dim wsTr As Worksheet
    Set wsTr = ThisWorkbook.Sheets(SH_TRANS)

    Dim positions As Object
    Set positions = BuildPositions(wsTr)

    Dim exRate As Double
    exRate = GetExRate(ThisWorkbook.Sheets(SH_PORT))

    Dim posData() As Variant
    Dim totalMktTWD As Double, totalCostTWD As Double, totalUnrlTWD As Double
    Dim posCount As Long
    Call CalcPositions(positions, exRate, posData, totalMktTWD, totalCostTWD, totalUnrlTWD, posCount)

    If Not IsArray(posData) Then Call NavNotify("V!: no positions found", True): Exit Sub
    Dim n As Long
    On Error Resume Next
    n = UBound(posData, 1)
    If Err.Number <> 0 Or n < 1 Then Call NavNotify("V!: no positions found", True): Exit Sub
    On Error GoTo 0

    Application.ScreenUpdating = False

    ' ---- 1. Fetch each holding's price history once, reuse for both steps ----
    Dim tickers() As String, mktVals() As Double, priceData() As Object
    ReDim tickers(0 To n - 1)
    ReDim mktVals(0 To n - 1)
    ReDim priceData(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        tickers(i) = CStr(posData(i + 1, 1))
        mktVals(i) = posData(i + 1, 8)
        Application.StatusBar = "Fetching [" & (i + 1) & "/" & n & "] " & tickers(i)
        DoEvents
        Set priceData(i) = GetHistoryPrices(tickers(i))
    Next i

    ' ---- 2. Per-stock 180D std dev, each on its own most-recent window ----
    Dim stockAnnVol() As Double, stockDaysUsed() As Long
    ReDim stockAnnVol(0 To n - 1)
    ReDim stockDaysUsed(0 To n - 1)

    Dim wantPx As Long: wantPx = 181   ' 181 closes -> 180 daily returns
    For i = 0 To n - 1
        Dim ownDates() As Date
        ownDates = VolSortedDateKeys(priceData(i))
        Dim ownCount As Long: ownCount = UBound(ownDates) - LBound(ownDates) + 1
        Dim useN As Long: useN = IIf(ownCount > wantPx, wantPx, ownCount)

        If useN >= 2 Then
            Dim pr() As Double
            ReDim pr(0 To useN - 1)
            Dim startIdx As Long: startIdx = ownCount - useN
            Dim k As Long
            For k = 0 To useN - 1
                pr(k) = priceData(i)(ownDates(startIdx + k))
            Next k
            Dim rtn() As Double
            rtn = VolDailyReturns(pr)
            stockDaysUsed(i) = UBound(rtn) - LBound(rtn) + 1
            If stockDaysUsed(i) >= 2 Then
                On Error Resume Next
                stockAnnVol(i) = Application.WorksheetFunction.StDev_S(rtn) * Sqr(252)
                On Error GoTo 0
            End If
        End If
    Next i

    ' ---- 3. Portfolio market-value-weighted std dev over the COMMON window ----
    Dim commonDates() As Date, usedCount As Long
    Call CorrGetCommonDates(priceData, n, commonDates, usedCount)

    Dim portAnnVol As Double, portDaysUsed As Long
    Dim retCount As Long: retCount = usedCount - 1
    If retCount >= 2 Then
        Dim portRet() As Double
        ReDim portRet(0 To retCount - 1)

        Dim j As Long
        For j = 0 To retCount - 1
            Dim wsum As Double: wsum = 0
            For i = 0 To n - 1
                Dim p0 As Double: p0 = 0
                Dim p1 As Double: p1 = 0
                On Error Resume Next
                p0 = priceData(i)(commonDates(j))
                p1 = priceData(i)(commonDates(j + 1))
                On Error GoTo 0
                If p0 > 0 And totalMktTWD > 0 Then
                    wsum = wsum + (mktVals(i) / totalMktTWD) * (p1 - p0) / p0
                End If
            Next i
            portRet(j) = wsum
        Next j

        portDaysUsed = retCount
        On Error Resume Next
        portAnnVol = Application.WorksheetFunction.StDev_S(portRet) * Sqr(252)
        On Error GoTo 0
    End If

    ' ---- 4. Render ----
    Call VolRenderSheet(tickers, mktVals, totalMktTWD, n, stockAnnVol, stockDaysUsed, _
                         portAnnVol, portDaysUsed, commonDates, usedCount)

    Application.ScreenUpdating = True
    Application.StatusBar = "Volatility updated: " & Format(Now, "hh:mm:ss")
    Call NavNotify("V! done " & Format(Now, "hh:mm:ss") & " - 180D volatility updated")
End Sub

' ------------------------------------------------------------
Private Function VolSortedDateKeys(dict As Object) As Date()
    Dim arr() As Date
    Dim cnt As Long: cnt = dict.Count
    If cnt <= 0 Then
        ReDim arr(0 To -1)
        VolSortedDateKeys = arr
        Exit Function
    End If

    Dim keys As Variant: keys = dict.Keys
    ReDim arr(0 To cnt - 1)
    Dim i As Long
    For i = 0 To cnt - 1
        arr(i) = CDate(keys(i))
    Next i

    Dim a As Long, b As Long, tmp As Date
    For a = 0 To cnt - 2
        For b = a + 1 To cnt - 1
            If arr(b) < arr(a) Then
                tmp = arr(a): arr(a) = arr(b): arr(b) = tmp
            End If
        Next b
    Next a
    VolSortedDateKeys = arr
End Function

' ------------------------------------------------------------
' Simple (arithmetic) daily returns - NOT log returns, see the comment
' above UpdatePortfolioVolatility for why simple returns are required
' for the market-value-weighted portfolio step.
Private Function VolDailyReturns(prices() As Double) As Double()
    Dim result() As Double
    Dim lo As Long: lo = LBound(prices)
    Dim n As Long: n = UBound(prices) - lo + 1
    If n < 2 Then
        ReDim result(0 To -1)
        VolDailyReturns = result
        Exit Function
    End If

    ReDim result(0 To n - 2)
    Dim i As Long
    For i = lo + 1 To UBound(prices)
        If prices(i - 1) > 0 Then
            result(i - 1 - lo) = (prices(i) - prices(i - 1)) / prices(i - 1)
        End If
    Next i
    VolDailyReturns = result
End Function

' ------------------------------------------------------------
Private Sub VolRenderSheet(tickers() As String, mktVals() As Double, totalMktTWD As Double, _
                            n As Long, stockAnnVol() As Double, stockDaysUsed() As Long, _
                            portAnnVol As Double, portDaysUsed As Long, _
                            commonDates() As Date, usedCount As Long)
    Dim wsV As Worksheet
    On Error Resume Next
    Set wsV = ThisWorkbook.Sheets("Volatility180D")
    On Error GoTo 0
    If wsV Is Nothing Then
        Set wsV = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsV.Name = "Volatility180D"
    End If

    Call NavStrip(wsV)
    wsV.Cells.Clear
    With wsV.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 10
    End With
    wsV.Activate
    ActiveWindow.DisplayGridlines = False

    With wsV.Cells(1, 1)
        .Value = "PORTFOLIO 180D VOLATILITY  |  Updated: " & Format(Now, "yyyy/mm/dd hh:mm:ss")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 14
    End With
    With wsV.Range(wsV.Cells(2, 1), wsV.Cells(2, 6))
        .Interior.Color = RGB(255, 192, 0)
        .RowHeight = 3
    End With

    Dim r As Long: r = 4
    wsV.Cells(r, 1).Value = "PORTFOLIO (MARKET-VALUE WEIGHTED)"
    wsV.Cells(r, 1).Font.Color = RGB(255, 192, 0)
    wsV.Cells(r, 1).Font.Bold = True
    r = r + 1

    Call WriteKV(wsV, r, "Annualized Std Dev", portAnnVol, "0.00%", False): r = r + 1
    Dim portDailyVol As Double: If portAnnVol > 0 Then portDailyVol = portAnnVol / Sqr(252)
    Call WriteKV(wsV, r, "Daily Std Dev", portDailyVol, "0.00%", False): r = r + 1

    wsV.Cells(r, 1).Value = "Trading Days Used"
    wsV.Cells(r, 1).Font.Color = RGB(150, 150, 150)
    wsV.Cells(r, 2).Value = portDaysUsed & " / 180"
    wsV.Cells(r, 2).Font.Color = IIf(portDaysUsed < 180, RGB(255, 192, 0), RGB(221, 221, 221))
    wsV.Cells(r, 2).Font.Bold = True
    r = r + 1

    If usedCount >= 2 Then
        wsV.Cells(r, 1).Value = "Common Date Range"
        wsV.Cells(r, 1).Font.Color = RGB(150, 150, 150)
        wsV.Cells(r, 2).Value = Format(commonDates(0), "yyyy/m/d") & "  ~  " & _
                                 Format(commonDates(usedCount - 1), "yyyy/m/d")
        wsV.Cells(r, 2).Font.Color = RGB(221, 221, 221)
        r = r + 1
    End If

    If portDaysUsed > 0 And portDaysUsed < 180 Then
        wsV.Cells(r, 1).Value = "* shortest-history holding capped the common window below 180 days"
        wsV.Cells(r, 1).Font.Color = RGB(150, 150, 150)
        wsV.Cells(r, 1).Font.Italic = True
        r = r + 1
    End If

    r = r + 2
    wsV.Cells(r, 1).Value = "PER-STOCK 180D STANDARD DEVIATION"
    wsV.Cells(r, 1).Font.Color = RGB(255, 192, 0)
    wsV.Cells(r, 1).Font.Bold = True
    r = r + 1

    Dim hdrs As Variant
    hdrs = Array("TICKER", "WEIGHT%", "DAYS USED", "DAILY STDEV", "ANNUALIZED STDEV")
    Dim c As Long
    For c = 0 To UBound(hdrs)
        With wsV.Cells(r, c + 1)
            .Value = hdrs(c)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next c
    r = r + 1

    Dim i As Long
    For i = 0 To n - 1
        Dim rowBg As Long: rowBg = IIf(i Mod 2 = 0, RGB(15, 15, 15), RGB(22, 22, 22))
        With wsV.Range(wsV.Cells(r, 1), wsV.Cells(r, 5))
            .Interior.Color = rowBg
            .HorizontalAlignment = xlCenter
        End With

        wsV.Cells(r, 1).Value = tickers(i)
        wsV.Cells(r, 1).Font.Color = RGB(255, 192, 0)
        wsV.Cells(r, 1).Font.Bold = True

        If totalMktTWD > 0 Then
            wsV.Cells(r, 2).Value = mktVals(i) / totalMktTWD
            wsV.Cells(r, 2).NumberFormat = "0.00%"
        End If

        wsV.Cells(r, 3).Value = stockDaysUsed(i) & IIf(stockDaysUsed(i) < 180, "*", "")
        wsV.Cells(r, 3).Font.Color = IIf(stockDaysUsed(i) < 180, RGB(255, 192, 0), RGB(200, 200, 200))

        Dim dailyVol As Double: If stockAnnVol(i) > 0 Then dailyVol = stockAnnVol(i) / Sqr(252)
        wsV.Cells(r, 4).Value = dailyVol
        wsV.Cells(r, 4).NumberFormat = "0.00%"

        wsV.Cells(r, 5).Value = stockAnnVol(i)
        wsV.Cells(r, 5).NumberFormat = "0.00%"
        wsV.Cells(r, 5).Font.Bold = True

        r = r + 1
    Next i

    r = r + 1
    wsV.Cells(r, 1).Value = "* fewer than 180 trading days of history available - used actual days shown"
    wsV.Cells(r, 1).Font.Color = RGB(150, 150, 150)
    wsV.Cells(r, 1).Font.Italic = True

    wsV.Columns("A:E").AutoFit
    Call NavAdd(wsV, "V")
    ThisWorkbook.Sheets(SH_PORT).Activate
End Sub


