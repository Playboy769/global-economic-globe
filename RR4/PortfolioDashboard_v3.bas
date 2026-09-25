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
'  Chart band       rows 26-39 weight donut at F (RR4_DONUT), a beta-
'                              exposure funnel across I:K (RR4_FUNNEL,
'                              v4.18) and the realized-PnL line across
'                              L:S (RR4_RLPNL, narrowed from J:S in v4.18)
'                              (v4.7 - both used to sit right of the panel)
'  (row 40 blank)
'  Position log                title row 41, headers row 42, data from 43
'  (row numbers in the notes below are the v4 ones; add RR4_TOP = 3)
'
'  Position log columns (BROKER column and broker group rows removed):
'    A TICKER   B NAME      C ENTRY DT  D DAYS      E SECTOR   F NET EXPOS
'    G SHARES   H ENTRY PX  I LAST      J % CHG     K UNRL PNL L WT%
'    M W.BETA   N BETA 30D  O P.TARGET  P SWING RISK  Q UPSIDE  R DOWNSIDE
'    (2026-09-25: UPSIDE / DOWNSIDE are hand-typed %, 25 = 25%; P.TARGET =
'    ENTRY PX x (1+UPSIDE%), SWING RISK = ENTRY PX x (1-DOWNSIDE%), formulas;
'    the Transactions P_Target column is no longer read.  Letters above are
'    the historical page-relative ones - real columns are P/Q/R/S.)
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
' v4.16 (2026-09-13): USD/TWD and ARRANGE <GO> labels sit ABOVE their input
' cells (row RR4_TOP+1, orange) instead of left of them; the inputs did not
' move (RR4_FX_CELL C6 / RR4_ARR_CELL E6), so no reader changed.
' v4.13 (2026-09-13): row 1 is now a blank spacer above the bar (like column
' A), so the bar sits in rows 2-4 and RR4_TOP went 3 -> 4; every absolute
' row constant below moved down by one and the config cells T1/T2 -> T2/T3.
' MigrateRR4TopRow inserts that row once on a sheet still on the v4.12 layout.
Public Const RR4_TOP       As Long = 4
' v4.2 (2026-09-12): column A is a blank spacer for breathing room, so the
' page also moved one column right - body columns are B:Q, the ticker panel
' J:S, the hidden order key V, and the config cells T1/T2 (they were S1/S2,
' which the panel now covers; RebuildPortfolioDashboard migrates them).
Public Const RR4_LEFT      As Long = 1
Public Const RR4_CFG_INC   As String = "T2"
Public Const RR4_CFG_CAP   As String = "T3"
' (v4.15, 2026-09-13: the big total that sat in B5 is gone - NET EXPOSURE in
'  the summary is the same number; page row 1 is blank at normal height)
Public Const RR4_FX_CELL   As String = "C6"
Public Const RR4_ARR_CELL  As String = "E6"
' 2026-09-23: live-refresh heartbeat cell (see ShowLiveRefreshStatus).
' NOT in the nav bar (rows 1-4) - the sheet has no freeze panes
' (COM-checked FreezePanes=False) and the user's normal scroll position
' sits around row 25+ to see the position table, so anything in rows 1-4
' is scrolled off-screen during actual use; that's why the first two
' placements (P2, then N2) never appeared for them even though the
' mechanism itself worked (screenshot-verified both times). L42 sits on
' the "POSITION LOG - RR4" title row, right above the table the user is
' already looking at - user-specified location, confirmed empty via COM.
Private Const LIVE_STATUS_ROW As Long = 42
Private Const LIVE_STATUS_COL As Long = 12   ' column L
' v4.7 (2026-09-12): a 14-row chart band (RR4_CHART_TOP..) sits between the
' upper blocks and the position log, which moved down from 27/28/29.
' v4.7.1: one blank row above (25) and below (40) the band; the ticker
' panel was shortened to row 24 to make room (TickerInsight TI_BOTTOM).
Public Const RR4_CHART_TOP As Long = 27
Public Const RR4_CHART_ROWS As Long = 14
' WATCHLIST (v4.9, 2026-09-12): B:E of the chart band, left of the donut.
' 2026-09-25: now a READ-ONLY SUMMARY of the Watch worksheet (modWatch,
' table tblWatch on WatchData): title row 26, header 27, row 28 unused
' (the old entry row), rows 29-35 = the first 7 names in Watch-page order,
' E = live last price (also written back to tblWatch's LAST cache), a row
' whose last <= target is lit. Double-click a row = open the Watch page on
' that name (SheetRR4_Code.txt -> modWatch.WatchGoto). ReadWatchlist /
' DrawWatchlist keep the block alive across the page clear.
Public Const RR4_WL_TITLE  As Long = 26
Public Const RR4_WL_HDR    As Long = 27
Public Const RR4_WL_ENTRY  As Long = 28
Public Const RR4_WL_FIRST  As Long = 29
Public Const RR4_WL_LAST   As Long = 35    ' 2026-09-21: 11 -> 7 rows to make room for TO-DO
' TO-DO (2026-09-21): B:E under the WATCHLIST. Row 36 = title + column
' names (TASK / DUE / DTE), row 37 = entry row (B ticker, C task, D due),
' rows 38-39 = the saved list, sorted by due date (undated last), DTE =
' due - today recomputed on every UP, lit orange once overdue.
' Double-click a saved row = done (TodoDeleteRow).
Public Const RR4_TD_TITLE  As Long = 36
Public Const RR4_TD_ENTRY  As Long = 37
Public Const RR4_TD_FIRST  As Long = 38
Public Const RR4_TD_LAST   As Long = 39
Public Const RR4_POS_TITLE As Long = 42
Public Const RR4_POS_HDR   As Long = 43
Public Const RR4_POS_FIRST As Long = 44
Private Const RR4_DONUT_NAME As String = "RR4_DONUT"
Private Const RR4_RLPNL_NAME As String = "RR4_RLPNL"   ' realized-PnL line chart (v4.6)
Private Const RR4_FUNNEL_NAME As String = "RR4_FUNNEL" ' beta-exposure funnel chart (v4.18)
Private Const RR4_FUN_PREFIX  As String = "RR4_FUN_"    ' shape-based funnel bars (see DrawFunnel)
Private Const RR4_NCOL      As Long = 19    ' last body column, B:S (R = UPSIDE, S = DOWNSIDE)
Private Const RR4_ORD_COL   As Long = 22    ' V (hidden)
Private Const RR4_SWING_COL As Long = 17    ' Q
Public Const RR4_UP_COL    As Long = 18    ' R  UPSIDE, hand-typed % (25 = 25%)
Public Const RR4_DN_COL    As Long = 19    ' S  DOWNSIDE, hand-typed % (12 or -12 = -12%)
Private Const RR4_UP_FG     As Long = 7237375   ' RGB(255,110,110)
Private Const RR4_DN_FG     As Long = 8570990   ' RGB(110,200,130)
Private Const RR4_LOG_ROWS  As Long = 5     ' daily-log trade lines, rows 10-14
Private Const RR4_POS_ROW_H As Double = 24  ' position-log data rows (v4.5, was 18)
Private Const RR4_WBAR_PREFIX As String = "RR4W_"
' Every cell the user types into is painted this dark grey (RGB 40,40,40)
' with WHITE text (RR4_INPUT_FG): the nav command cell C1, USD/TWD C5,
' ARRANGE E5, SWING RISK (Q29:Q..), the ticker panel's K4 / L20, and the
' B2 input of the VT / CC pages. (v4.3, 2026-09-12: was 70,70,70 + yellow.)
Public Const RR4_INPUT_BG  As Long = 2631720
Public Const RR4_INPUT_FG  As Long = 16777215
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
Private m_fxLive As Boolean                 ' v4.14: was the last UP's USD/TWD fetched live (Yahoo TWD=X)
' Per-SELL FIFO detail recorded by BuildPositions, keyed by Transactions sheet
' row: Array(firstEntrySer, lastEntrySer, shareWeightedDays, costBasis).
' DrawDailyLog reads it so an EXIT/TRIM line can show what was held and for how
' long without walking FIFO a second time (the three FIFO walkers must agree).
Private m_sellInfo As Object

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
    Call MigrateRR4TopRow(wsP)          ' v4.13: one-time shift onto the blank-row-1 layout (before any read)
    Call MigrateTransactionsNav         ' 2026-09-13: Transactions carries the bar too (no-op once it does)
    Call MigrateHistoryLogNav           ' 2026-09-13: and HistoryLog (no-op once it does)
    Call DrawTransactionsHeader(wsTr)   ' header in the RR4 palette + AutoFilter over the table

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
    ' USD/TWD (v4.14, 2026-09-13): live from Yahoo TWD=X on every UP; the
    ' typed C6 value only carries the rate across when the fetch fails
    Dim exRate As Double: exRate = GetExRate(wsP)
    Dim liveFx As Double: liveFx = FetchLiveFx()
    m_fxLive = (liveFx > 20 And liveFx < 50)
    If m_fxLive Then exRate = liveFx
    Dim arrCode As String: arrCode = UCase(CellStr(wsP.Range(RR4_ARR_CELL).Value))
    Dim tiTicker As String: tiTicker = UCase(CellStr(wsP.Range(TI_TICKER_CELL).Value))
    If tiTicker = "" Then tiTicker = UCase(CellStr(wsP.Range("K5").Value))   ' one-time: pre-v2.5 panel kept it in K5
    Dim tiTarget As Variant: tiTarget = wsP.Range(TI_TARGET_CELL).Value
    If IsEmpty(tiTarget) Then tiTarget = wsP.Range("L21").Value   ' one-time: pre-v2.6 panel kept it in L21
    ' config (SetupPortfolioConfig): T1/T2 since v4.2, S1/S2 before it - the
    ' panel covers S now, so carry the old pair over once
    Dim cfgInc As Variant: cfgInc = wsP.Range(RR4_CFG_INC).Value
    Dim cfgCap As Variant: cfgCap = wsP.Range(RR4_CFG_CAP).Value
    If Not IsDate(cfgInc) Then cfgInc = wsP.Range("S1").Value
    If NumOr0(cfgCap) <= 0 Then cfgCap = wsP.Range("S2").Value
    Dim upMap As Object: Set upMap = ReadHandColumn(wsP, "UPSIDE")
    Dim dnMap As Object: Set dnMap = ReadHandColumn(wsP, "DOWNSIDE")
    Dim wl As Variant: wl = ReadWatchlist(wsP)
    Dim td As Variant: td = ReadTodo(wsP)
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

    ' 2026-09-23: 一次 UP 只發 1 次 MIS 批次請求，涵蓋持倉+watchlist+ticker
    ' panel 全部台股代號，避免每檔各查一次撞到 MIS 的節流門檻。GetStockPrice
    ' 內部會在這之後的每一次呼叫優先讀這批 prefetch 的快取。
    Dim twTickers() As String
    twTickers = CollectTWTickersForPrefetch(positions, wl, tiTicker)
    modMISPrice.PrefetchMISPrices twTickers

    Dim posData()    As Variant
    Dim totalMktTWD  As Double
    Dim totalCostTWD As Double
    Dim totalUnrlTWD As Double
    Dim posCount     As Long

    Call CalcPositions(positions, exRate, posData, totalMktTWD, totalCostTWD, totalUnrlTWD, posCount)
    Call CalculateRealizedPnL

    Dim realPnL As Double
    On Error Resume Next
    ' PNL(TWD) moved G -> H when RET% was inserted as column B (Attach.bas);
    ' page column 8, wherever the nav bar puts it (Attach.RealCol)
    realPnL = Application.WorksheetFunction.Sum(wsR.Columns(RealCol(wsR, 8)))
    On Error GoTo Fail

    Dim portBeta As Double
    portBeta = CalcPortBeta(posData)

    ' previous trading day's cumulative PnL - read before LogHistory
    ' rewrites / appends today's HistoryLog row
    Dim prevPnL As Variant: prevPnL = PrevDayCumPnL()

    Call DrawHeader(wsP, exRate, arrCode)
    Call DrawDailyLog(wsP, posData, exRate, totalMktTWD, totalUnrlTWD + realPnL, realPnL, prevPnL)
    Call DrawSummary(wsP, posData, totalMktTWD, totalCostTWD, totalUnrlTWD, realPnL, portBeta, posCount)
    Call DrawColumnHeaders(wsP)
    Dim lastDataRow As Long
    lastDataRow = WritePositionRows(wsP, posData, totalMktTWD, upMap, dnMap)
    Call ApplyArrange(arrCode)
    Call DrawWatchlist(wsP, wl)
    Call DrawTodo(wsP, td)
    Call DrawDisclaimer(wsP, lastDataRow)
    Call RenderTickerPanel(tiTicker, tiTarget)
    Call LogHistory(totalMktTWD, totalUnrlTWD + realPnL, realPnL)
    Call DrawRealizedChart(wsP)     ' after LogHistory, so today's row is on the line

    Application.ScreenUpdating = True
    ' (v4.6: the separate "Analysis" sheet / DrawDeepAnalysis is gone - its
    ' currency split lives in the Summary, the rest was not used)
    Application.EnableEvents = prevEvents
    Application.StatusBar = "Dashboard updated: " & Format(Now, "hh:mm:ss")
    Call NavNotify("UP done " & Format(Now, "hh:mm:ss") & " - " & posCount & " positions  .  USD/TWD " & Format(exRate, "0.00") & IIf(m_fxLive, " live", " (typed / default - Yahoo fetch failed)"))
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

' PAGE columns (A=1); the sheet column is HistCol(wsH, n) - the page
' carries the nav bar since 2026-09-13 (blank row 1 / column A, bar 2-4).
Private Function HistPriceCols() As Variant
    HistPriceCols = Array(6, 8, 10, 12)          ' F H J L
End Function

Private Function HistRetCols() As Variant
    HistRetCols = Array(7, 9, 11, 13)            ' G I K M
End Function

' One-off move of HistoryLog onto the nav-bar layout. NavAdd inserts the
' blank row / column and the bar; before that the leftovers the owner
' chose to drop go: the #DIV/0! "(O-$O$2)/$O$2" formulas + one stranded
' note in page columns Q:R, the hand-typed X1:X2 pair, and the two "=0"
' conditional formats on H:L. Every formula row is then rewritten by
' WriteHistoryRowFormulas so the sheet matches what LogHistory produces.
' No-op once the bar is there.
Private Sub MigrateHistoryLogNav()
    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = ThisWorkbook.Sheets(SH_HIST): On Error GoTo 0
    If wsH Is Nothing Then Exit Sub
    If NavHasRows(wsH) Then Exit Sub
    On Error Resume Next
    wsH.Cells.FormatConditions.Delete
    wsH.Range(wsH.Columns(HIST_NCOL + 1), wsH.Columns(HIST_NCOL + 13)).Clear     ' page N:Z
    On Error GoTo 0
    Call NavAdd(wsH, "H")
    Dim r As Long
    For r = HistHdrRow(wsH) + 1 To HistLastRow(wsH)
        Call WriteHistoryRowFormulas(wsH, r)
    Next r
End Sub

Private Sub LogHistory(totalMkt As Double, totalPnL As Double, realPnL As Double)
    Dim wsH As Worksheet
    On Error Resume Next
    Set wsH = ThisWorkbook.Sheets(SH_HIST)
    On Error GoTo 0
    If wsH Is Nothing Then Exit Sub

    Call EnsureHistoryHeaders(wsH)
    Call EnsureYTDBaselines(False)

    Dim hc As Long: hc = NavLeft(wsH)              ' page column n -> sheet column hc + n
    Dim nr As Long: nr = HistLastRow(wsH) + 1
    If nr > HistHdrRow(wsH) + 1 Then
        ' guard the date: text or an error value in column A used to abort the
        ' whole rebuild here, after the page had already been drawn
        Dim lastStamp As Variant: lastStamp = wsH.cells(nr - 1, hc + 1).Value
        If IsDate(lastStamp) Then
            If Int(CDate(lastStamp)) = Date Then nr = nr - 1
        End If
    End If

    With wsH
        .cells(nr, hc + 1).Value = Now
        .cells(nr, hc + 2).Value = totalMkt
        .cells(nr, hc + 3).Value = totalPnL
        .cells(nr, hc + 4).Value = realPnL

        .cells(nr, hc + 1).NumberFormat = "yyyy/m/d h:mm:ss"
        .Range(.cells(nr, hc + 2), .cells(nr, hc + 4)).NumberFormat = "#,##0"
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
            wsH.cells(nr, hc + priceCols(i)).Value = px
        Else
            ' download failed - leave the cell blank instead of logging a 0
            wsH.cells(nr, hc + priceCols(i)).ClearContents
        End If
        wsH.cells(nr, hc + priceCols(i)).NumberFormat = "#,##0.00"
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

    ' realized trades: (exit date, PnL TWD) - page columns I / H via Attach.RealCol
    Dim r0 As Long: r0 = RealHdrRow(wsR)
    Dim lastR As Long: lastR = RealLastRow(wsR)
    Dim n As Long, r As Long
    Dim exD() As Long, exP() As Double
    If lastR > r0 Then
        ReDim exD(1 To lastR - r0): ReDim exP(1 To lastR - r0)
        For r = r0 + 1 To lastR
            Dim dv As Variant: dv = wsR.cells(r, RealCol(wsR, 9)).Value
            If IsDate(dv) Then
                n = n + 1
                exD(n) = Int(CDbl(CDate(dv)))
                exP(n) = NumOr0(wsR.cells(r, RealCol(wsR, 8)).Value)
            End If
        Next r
    End If

    Dim lastH As Long: lastH = HistLastRow(wsH)
    For r = HistHdrRow(wsH) + 1 To lastH
        Dim hv As Variant: hv = wsH.cells(r, HistCol(wsH, 1)).Value
        If IsDate(hv) Then
            Dim dayN As Long: dayN = Int(CDbl(CDate(hv)))
            Dim cum As Double: cum = 0
            Dim k As Long
            For k = 1 To n
                If exD(k) <= dayN Then cum = cum + exP(k)
            Next k
            If Abs(NumOr0(wsH.cells(r, HistCol(wsH, 4)).Value) - cum) > 0.005 Then wsH.cells(r, HistCol(wsH, 4)).Value = cum
        End If
    Next r
End Sub

' Header row - rewritten every run so the layout stays self-describing
' (page A1:M1; RR4 palette since 2026-09-13, and the nav bar repainted).
Private Sub EnsureHistoryHeaders(wsH As Worksheet)
    Call NavAdd(wsH, "H")
    Dim r0 As Long: r0 = HistHdrRow(wsH)
    Dim hdr As Variant
    hdr = Array("Date", "TotalMarketValue", "TotalCumulativePnL", _
                "Realized PnL", "Realized PnL Daily Chg%", _
                "Price (SPY)", "YTD Ret% (SPY)", _
                "Price (QQQ)", "YTD Ret% (QQQ)", _
                "Price (TWII)", "YTD Ret% (TWII)", _
                "Price (SOX)", "YTD Ret% (SOX)")
    Dim i As Long
    For i = LBound(hdr) To UBound(hdr)
        With wsH.cells(r0, HistCol(wsH, i + 1))
            If CStr(.Value) <> CStr(hdr(i)) Then .Value = hdr(i)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .Font.Size = 9
            .Font.Name = "Consolas"
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next i
    With wsH.Range(wsH.cells(r0, HistCol(wsH, 1)), wsH.cells(r0, HistCol(wsH, HIST_NCOL))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With
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

    Dim i As Long, pc As String, rc As Long, bc As String
    For i = LBound(retCols) To UBound(retCols)
        pc = wsH.cells(r, HistCol(wsH, priceCols(i))).Address(False, False)   ' sheet address of the page cell
        rc = HistCol(wsH, retCols(i))
        ' each index divides by its OWN baseline on the raw-data sheet
        bc = SH_HRAW & "!$C$" & (HL_BASE_ROW + i)
        wsH.cells(r, rc).Formula = "=IFERROR(IF(" & pc & "="""","""",(" & pc & "-" & bc & ")/" & bc & "),"""")"
        wsH.cells(r, rc).NumberFormat = "0.00%"
    Next i

    ' Realized PnL daily change - "-" instead of #DIV/0! while realized PnL is 0
    Dim eCol As Long: eCol = HistCol(wsH, 5)
    If r > HistHdrRow(wsH) + 1 Then
        Dim dNow As String: dNow = wsH.cells(r, HistCol(wsH, 4)).Address(False, False)
        Dim dPrev As String: dPrev = wsH.cells(r - 1, HistCol(wsH, 4)).Address(False, False)
        wsH.cells(r, eCol).Formula = "=IFERROR((" & dNow & "-" & dPrev & ")/ABS(" & dPrev & "),""-"")"
    Else
        wsH.cells(r, eCol).Formula = "=""-"""
    End If
    wsH.cells(r, eCol).NumberFormat = "0.00%"
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

    Call MigrateHistoryLogNav
    Dim lastRow As Long
    lastRow = HistLastRow(wsH)
    If lastRow <= HistHdrRow(wsH) Then MsgBox "No data to backfill": Exit Sub

    Call EnsureHistoryHeaders(wsH)
    Call EnsureYTDBaselines(False)

    Application.ScreenUpdating = False

    Dim tickers As Variant, priceCols As Variant
    tickers = HistTickers()
    priceCols = HistPriceCols()
    Dim hc As Long: hc = NavLeft(wsH)

    Dim filled() As Long
    ReDim filled(LBound(tickers) To UBound(tickers))

    Dim i As Long, k As Long
    For i = HistHdrRow(wsH) + 1 To lastRow
        If wsH.cells(i, hc + 1).Value = "" Then GoTo NextRow

        Dim targetDate As Date
        targetDate = 0
        On Error Resume Next
        targetDate = CDate(Int(CDbl(wsH.cells(i, hc + 1).Value)))
        On Error GoTo 0
        If targetDate = 0 Then GoTo NextRow

        For k = LBound(tickers) To UBound(tickers)
            If wsH.cells(i, hc + priceCols(k)).Value = "" Then
                Dim price As Double
                price = GetHistoricalPrice(CStr(tickers(k)), targetDate)
                If price > 0 Then
                    wsH.cells(i, hc + priceCols(k)).Value = price
                    filled(k) = filled(k) + 1
                End If
            End If
            wsH.cells(i, hc + priceCols(k)).NumberFormat = "#,##0.00"
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

    Call MigrateHistoryLogNav
    Dim lastRow As Long
    lastRow = HistLastRow(wsH)
    Dim n As Long: n = lastRow - HistHdrRow(wsH)
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

    Dim hc As Long: hc = NavLeft(wsH)
    wsH.Range(wsH.cells(HistHdrRow(wsH), hc + 14), wsH.cells(lastRow, hc + 19)).ClearContents   ' page N:S

    Call EnsureHistoryHeaders(wsH)
    Call EnsureYTDBaselines(True)

    Dim tickers As Variant, priceCols As Variant
    tickers = HistTickers()
    priceCols = HistPriceCols()

    Dim missing As Long, r As Long, k As Long
    For r = HistHdrRow(wsH) + 1 To lastRow
        Dim targetDate As Date
        targetDate = 0
        On Error Resume Next
        targetDate = CDate(Int(CDbl(wsH.cells(r, hc + 1).Value)))
        On Error GoTo 0

        ' D currently holds whatever the broken run left there (old SPY
        ' prices on the pre-repair rows), and the real values are gone.
        wsH.cells(r, hc + 4).Value = 0
        wsH.Range(wsH.cells(r, hc + 2), wsH.cells(r, hc + 4)).NumberFormat = "#,##0"

        For k = LBound(tickers) To UBound(tickers)
            Dim price As Double: price = 0
            If targetDate > 0 Then price = GetHistoricalPrice(CStr(tickers(k)), targetDate)
            If price > 0 Then
                wsH.cells(r, hc + priceCols(k)).Value = price
            Else
                wsH.cells(r, hc + priceCols(k)).ClearContents
                missing = missing + 1
            End If
            wsH.cells(r, hc + priceCols(k)).NumberFormat = "#,##0.00"
        Next k

        Call WriteHistoryRowFormulas(wsH, r)
        Application.StatusBar = "Repairing row " & r & "/" & lastRow
    Next r

    wsH.Range(wsH.Columns(hc + 1), wsH.Columns(hc + HIST_NCOL)).EntireColumn.AutoFit
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
    ws.cells.UnMerge          ' the NOTE header (R:S) is re-merged by DrawColumnHeaders
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
    ' columns; P.TARGET / SWING RISK are formulas; R/S are hand-typed UPSIDE / DOWNSIDE %.
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
End Sub

' ================================================================
'  Header (page rows 1-4 = sheet rows 5-8): USD/TWD, ARRANGE <GO>,
'  weight-bar label. Page row 1 is blank (v4.15 dropped the big total).
' ================================================================
Private Sub DrawHeader(ws As Worksheet, exRate As Double, arrCode As String)
    ' v4.16 (2026-09-13): label ABOVE its input (page row 1), input cells stay C6 / E6
    Call StackedLabel(ws.Range(RR4_FX_CELL).Offset(-1, 0), "USD/TWD")
    With ws.Range(RR4_FX_CELL)
        .Value = exRate
        .NumberFormat = "0.00"
        ' v4.17 (2026-09-14): the rate is fetched live (FetchLiveFx), so the cell is
        ' no longer painted as an input - black like the page; typing still works offline
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        ' where the rate came from (hover): live Yahoo TWD=X, or the typed / default fallback
        On Error Resume Next
        .ClearComments
        .AddComment IIf(m_fxLive, "USD/TWD live from Yahoo TWD=X, " & Format(Now, "yyyy/mm/dd hh:mm"), _
                        "USD/TWD: Yahoo fetch failed - using the typed value (31.6 default)")
        On Error GoTo 0
    End With

    Call StackedLabel(ws.Range(RR4_ARR_CELL).Offset(-1, 0), "ARRANGE <GO>")
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

' Orange bold label sitting on the row above its input cell, left-aligned with it
' and allowed to overflow to the right (ShrinkToFit off).
Private Sub StackedLabel(cell As Range, ByVal txt As String)
    With cell
        .Value = txt
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
    End With
End Sub

' ================================================================
'  DAILY LOG (page rows 6-12 = sheet rows 9-15) - today only, not kept
' ----------------------------------------------------------------
'  One line per BUY / SELL dated today on Transactions:
'    NEW  = a buy into a ticker that held no shares before today
'    ADD  = a buy into an existing position
'    TRIM = a sell that leaves shares
'    EXIT = a sell that closes the position
'  TRIM / EXIT lines also carry "HELD in > out (Nd)   RET x%": the entry
'  date(s) of the FIFO lots the sell consumed, share-weighted holding days,
'  and the return on their acquisition cost (= the Realized page's RET%).
'  The detail comes from m_sellInfo, which BuildPositions fills - so this
'  must run after BuildPositions in the same UP (RebuildPortfolioDashboard
'  does).
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
    Dim lastR As Long: lastR = TrLastRow(wsTr)
    Dim tc As Long: tc = NavLeft(wsTr)             ' Transactions page column n -> tc + n
    Dim tRows() As Long: ReDim tRows(1 To IIf(lastR > 1, lastR, 1))
    Dim cnt As Long

    ' pass 1: today's rows, and today's net shares per ticker.
    ' Every fill is its own row (the 2026-09-15 monthly Buy consolidation was
    ' withdrawn on 2026-09-18), so the Date column alone identifies them.
    Dim todayNet As Object: Set todayNet = CreateObject("Scripting.Dictionary")
    Dim r As Long, act As String, tk As String, sh As Double
    For r = TrHdrRow(wsTr) + 1 To lastR
        Dim dv As Variant: dv = wsTr.cells(r, tc + 2).Value
        If IsDate(dv) Then
            If Int(CDate(dv)) = Date Then
                act = UCase(CellStr(wsTr.cells(r, tc + 4).Value))
                If act = "BUY" Or act = "SELL" Then
                    cnt = cnt + 1
                    tRows(cnt) = r
                    tk = UCase(CellStr(wsTr.cells(r, tc + 3).Value))
                    sh = NumOr0(wsTr.cells(r, tc + 5).Value)
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
        act = UCase(CellStr(wsTr.cells(r, tc + 4).Value))
        tk = UCase(CellStr(wsTr.cells(r, tc + 3).Value))
        sh = NumOr0(wsTr.cells(r, tc + 5).Value)
        Dim px As Double: px = NumOr0(wsTr.cells(r, tc + 6).Value)
        Dim amt As Double: amt = NumOr0(wsTr.cells(r, tc + 9).Value)
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

        ' sells: what was held, for how long, and the return - FIFO detail
        ' BuildPositions recorded for this sheet row (same rule as Realized RET%)
        Dim heldTxt As String: heldTxt = ""
        Dim retV As Variant: retV = Empty
        If act = "SELL" And Not m_sellInfo Is Nothing Then
            If m_sellInfo.Exists(r) Then
                Dim si As Variant: si = m_sellInfo(r)
                heldTxt = SellHeldText(CLng(si(0)), CLng(si(1)), CDbl(si(2)), wsTr.cells(r, tc + 2).Value)
                If si(3) > 0 Then retV = (amt - si(3)) / si(3)
            End If
        End If

        If cnt <= RR4_LOG_ROWS Or k < RR4_LOG_ROWS Then
            Call WriteLogLine(ws, outRow, lineTag, tk, sh, px, amtTWD, heldTxt, retV)
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

' held / retV are only passed for sells: "HELD 9/7 > 9/18 (11d)   RET +7.72%",
' RET% coloured like every other gain/loss on this page.
Private Sub WriteLogLine(ws As Worksheet, r As Long, lineTag As String, tk As String, _
                         sh As Double, px As Double, amtTWD As Double, _
                         Optional ByVal held As String = "", Optional ByVal retV As Variant)
    With ws.cells(r, RR4_LEFT + 1)
        .Value = lineTag
        .Font.Bold = True
        If lineTag = "NEW" Or lineTag = "ADD" Then
            .Font.Color = RR4_ACCENT
        Else
            .Font.Color = RGB(0, 200, 255)
        End If
    End With
    Dim txt As String
    txt = tk & "   " & FormatShares(sh) & " SH   @ " & Format(px, "#,##0.00") & _
          "   |   " & Format(amtTWD, "#,##0") & " TWD"
    Dim retTxt As String
    If held <> "" Then
        txt = txt & "   |   " & held
        If Not IsMissing(retV) Then
            If Not IsEmpty(retV) Then
                retTxt = Format(retV, "+0.00%;-0.00%;0.00%")
                txt = txt & "   RET " & retTxt
            End If
        End If
    End If
    txt = UCase(txt)          ' owner's choice: the whole log line in capitals
    With ws.cells(r, RR4_LEFT + 2)
        .Value = txt
        .Font.Color = RGB(221, 221, 221)
        If retTxt <> "" Then
            With .Characters(Len(txt) - Len(retTxt) + 1, Len(retTxt)).Font
                .Color = GainLossColor(CDbl(retV))
                .Bold = True
            End With
        End If
    End With
End Sub

' "HELD 9/7 > 9/18 (11D)" for a sell that consumed one lot; when it ate
' several lots bought on different days: "HELD 8/7 - 8/12 > 9/10 (AVG 29D)",
' days weighted by the shares taken from each lot. The entry date carries
' the year only when it differs from the exit's.
Private Function SellHeldText(ByVal firstIn As Long, ByVal lastIn As Long, _
                              ByVal avgDays As Double, ByVal exitV As Variant) As String
    If firstIn = 0 Then SellHeldText = "HELD ?": Exit Function
    Dim exitD As Date: If IsDate(exitV) Then exitD = CDate(exitV)
    Dim fmtIn As String: fmtIn = "m/d"
    If IsDate(exitV) Then
        If Year(CDate(firstIn)) <> Year(exitD) Then fmtIn = "yyyy/m/d"
    End If
    Dim s As String: s = "HELD " & Format(CDate(firstIn), fmtIn)
    If lastIn <> firstIn Then s = s & " - " & Format(CDate(lastIn), fmtIn)
    If IsDate(exitV) Then s = s & " > " & Format(exitD, "m/d")
    If lastIn <> firstIn Then
        s = s & " (AVG " & Format(avgDays, "0") & "D)"
    Else
        s = s & " (" & Format(avgDays, "0") & "D)"
    End If
    SellHeldText = s
End Function

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
    With ws.Range(ws.cells(RR4_POS_TITLE, RR4_LEFT + 1), ws.cells(RR4_POS_TITLE, RR4_NCOL)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    Dim headers As Variant
    headers = Array("TICKER", "NAME", "ENTRY DT", "DAYS", "SECTOR", _
                    "NET EXPOS", "SHARES", "ENTRY PX", "LAST", "% CHG", _
                    "UNRL PNL", "WT%", "W.BETA", "BETA 30D", "P.TARGET", "SWING RISK", "UPSIDE", "DOWNSIDE")
    ' the old NOTE header was merged over R:S; clear that before writing S
    ws.Range(ws.cells(RR4_POS_HDR, RR4_UP_COL), ws.cells(RR4_POS_HDR, RR4_DN_COL)).UnMerge
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
                                    totalMkt As Double, ByVal upMap As Object, _
                                    ByVal dnMap As Object) As Long
    WritePositionRows = RR4_POS_FIRST - 1
    Dim n As Long: n = PosCountOf(posData)
    If n < 1 Then Exit Function

    Dim i As Long, r As Long: r = RR4_POS_FIRST
    For i = 1 To n
        Call WriteOnePositionRow(ws, r, i, posData, totalMkt, upMap, dnMap)
        r = r + 1
    Next i
    WritePositionRows = r - 1
    ' the workbook may sit in manual calculation - Range.Calculate works either way
    ws.Range(ws.cells(RR4_POS_FIRST, RR4_LEFT + 15), ws.cells(r - 1, RR4_SWING_COL)).Calculate

    Call DrawTopExposure(ws, posData, totalMkt)
End Function

' Typing UPSIDE / DOWNSIDE: refresh that row's P.TARGET / SWING RISK formulas
' (called from the RR4 sheet's Worksheet_Change; manual-calculation safe).
Public Sub RecalcTargetRow(ByVal ws As Worksheet, ByVal r As Long)
    ws.Range(ws.cells(r, RR4_LEFT + 15), ws.cells(r, RR4_SWING_COL)).Calculate
End Sub

' ------------------------------------------------------------
Private Sub WriteOnePositionRow(ws As Worksheet, r As Long, i As Long, _
                                 posData() As Variant, totalMkt As Double, _
                                 upMap As Object, dnMap As Object)
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

    ' P.TARGET = ENTRY PX x (1 + UPSIDE%), SWING RISK = ENTRY PX x (1 - |DOWNSIDE|%):
    ' live formulas on the row's own R / S / I cells (so they follow a hand edit
    ' and survive ARRANGE sorts); blank until UPSIDE / DOWNSIDE is typed. The old
    ' Transactions P_Target column is no longer read.
    With ws.cells(r, RR4_LEFT + 15)
        .NumberFormat = "#,##0.00"
        .FormulaR1C1 = "=IF(AND(ISNUMBER(RC[2]),RC[-7]>0),RC[-7]*(1+RC[2]/100),"""")"
        .Font.Color = RR4_ACCENT
    End With
    With ws.cells(r, RR4_SWING_COL)
        .NumberFormat = "#,##0.00"
        .FormulaR1C1 = "=IF(AND(ISNUMBER(RC[2]),RC[-8]>0),RC[-8]*(1-ABS(RC[2])/100),"""")"
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
    End With
    ' UPSIDE / DOWNSIDE: plain numbers shown as "n%" by the format (no /100), so
    ' typing 25 reads +25.0%; the DOWNSIDE format drops the sign of the value and
    ' prints "-", so 12 and -12 both read -12.0%.
    With ws.cells(r, RR4_UP_COL)
        .NumberFormat = "+0.0""%"";-0.0""%"";0.0""%"""
        If upMap.Exists(tickerCode) Then .Value = HandNumber(upMap(tickerCode))
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
    End With
    With ws.cells(r, RR4_DN_COL)
        .NumberFormat = """-""0.0""%"";""-""0.0""%"";0.0""%"""
        If dnMap.Exists(tickerCode) Then .Value = HandNumber(dnMap(tickerCode))
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
    End With

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
        ' 2026-09-18 fix: these were hardcoded as page-relative offsets (as if
        ' RR4_LEFT were 0), one column short of the field named in each desc
        ' string - e.g. "UNU" sorted by % CHG (col 11) while claiming UNRL
        ' PNL (col 12). Rows did reorder, just by the wrong metric each time.
        ' Written as RR4_LEFT+n to match WriteOnePositionRow/the header array
        ' so a future row-shift can't silently reintroduce the same drift.
        Case "", "DEF": keyCol = RR4_ORD_COL: ord = xlAscending: desc = "DEFAULT (TW code, then US A-Z)"
        Case "UNU": keyCol = RR4_LEFT + 11: ord = xlAscending: desc = "UNRL PNL  low > high"
        Case "UND": keyCol = RR4_LEFT + 11: ord = xlDescending: desc = "UNRL PNL  high > low"
        Case "PCU": keyCol = RR4_LEFT + 10: ord = xlAscending: desc = "% CHG  low > high"
        Case "PCD": keyCol = RR4_LEFT + 10: ord = xlDescending: desc = "% CHG  high > low"
        Case "DAU": keyCol = RR4_LEFT + 4: ord = xlAscending: desc = "DAYS  short > long"
        Case "DAD": keyCol = RR4_LEFT + 4: ord = xlDescending: desc = "DAYS  long > short"
        Case "WTU": keyCol = RR4_LEFT + 12: ord = xlAscending: desc = "WT%  light > heavy"
        Case "WTD": keyCol = RR4_LEFT + 12: ord = xlDescending: desc = "WT%  heavy > light"
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
    Call DrawFunnel(ws, lastR)

    Dim codeList As String: codeList = "UNU/UND PCU/PCD DAU/DAD WTU/WTD DEF  >>  "
    With ws.cells(RR4_TOP + 2, RR4_LEFT + 5)
        .Value = codeList & desc
        .Font.Size = 9
        .Font.Color = RGB(0, 200, 255)
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
        ' v4.11: SWING RISK rides the stripe like every other cell; UPSIDE red,
        ' DOWNSIDE green (TW convention)
        ws.cells(r, RR4_SWING_COL).Font.Color = RR4_INPUT_FG
        ws.cells(r, RR4_UP_COL).Font.Color = RR4_UP_FG
        ws.cells(r, RR4_DN_COL).Font.Color = RR4_DN_FG
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

    ' 2026-09-21: snap every edge to whole screen pixels (0.75pt). With
    ' fractional tops/lefts Excel anti-aliases each rectangle differently,
    ' so identical-height boxes rendered visibly uneven.
    barT = PxSnap(barT): barH = PxSnap(barH)
    Dim x As Double: x = barL
    For r = RR4_POS_FIRST To lastR
        Dim wt As Double: wt = Abs(NumOr0(ws.cells(r, RR4_LEFT + 12).Value))
        Dim w As Double: w = wt / sumW * barW
        If w >= 1 Then
            Dim pct As Double: pct = NumOr0(ws.cells(r, RR4_LEFT + 10).Value)
            Dim tk As String: tk = CellStr(ws.cells(r, RR4_LEFT + 1).Value)
            Dim sx As Double: sx = PxSnap(x)
            Dim sw As Double: sw = PxSnap(x + w) - sx
            If sw > 1.5 Then sw = sw - 0.75       ' 1px gap between boxes
            If sw < 0.75 Then sw = 0.75
            Dim shp As Shape
            Set shp = ws.Shapes.AddShape(msoShapeRectangle, sx, barT, sw, barH)
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

' Round a point value to the nearest whole screen pixel (1px = 0.75pt).
Private Function PxSnap(ByVal pt As Double) As Double
    PxSnap = Round(pt / 0.75, 0) * 0.75
End Function

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
            .Font.Color = RGB(190, 190, 190)     ' light grey (v4.10.1)
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
'  BETA EXPOSURE FUNNEL (v4.18, 2026-09-19) - columns I:K of the chart
'  band, between the weight donut (F) and the realized-PnL line (now
'  narrowed to L:S to make room). One stage per position in the current
'  ARRANGE order (same rows the donut and weight bar read), width =
'  W.BETA (column RR4_LEFT+13, WT% x Beta - see WriteOnePositionRow) so
'  the funnel reads as each position's share of the portfolio's
'  systematic-risk exposure rather than raw dollar weight.
'
'  Drawn as shapes, not a real chart object - confirmed live via COM
'  2026-09-19 that Excel's native Funnel chart type (123) has NO working
'  VBA series API on this host (Excel 16.0 build 17932): ChartObjects.Add
'  + ChartType=123 raises "Value does not fall within the expected
'  range"; Shapes.AddChart2(-1, 123, ...) *does* create the chart object,
'  but Series.Values/.XValues (Range OR literal array) silently collapse
'  every position down to a single point, and Chart.SetSourceData throws
'  "The method or operation is not implemented". This is a documented
'  Microsoft limitation of the Funnel chart type's object model, not a
'  fixable call-order bug - Waterfall/Treemap have real VBA support,
'  Funnel does not. So this reads like the fake doughnut-hole trick in
'  DrawDonut: one trapezoid shape per position (msoShapeTrapezoid, named
'  RR4_FUN_n so only they are cleared on redraw), width scaled to
'  |W.BETA| / max(|W.BETA|) and centred.
'  2026-09-20: each trapezoid is filled with DonutColor(i) using the same
'  row index the WEIGHT donut reads, so a given position is the same
'  colour in both charts (was solid RR4_ACCENT for every stage).
' ================================================================
Private Sub DrawFunnel(ws As Worksheet, lastR As Long)
    Dim k As Long
    ' one-time cleanup: an earlier build of this sub created a (non-functional)
    ' native Funnel ChartObject under this name - remove it if still present
    For k = ws.ChartObjects.count To 1 Step -1
        If ws.ChartObjects(k).Name = RR4_FUNNEL_NAME Then ws.ChartObjects(k).Delete
    Next k
    For k = ws.Shapes.count To 1 Step -1
        If Left(ws.Shapes(k).Name, Len(RR4_FUN_PREFIX)) = RR4_FUN_PREFIX Then ws.Shapes(k).Delete
    Next k
    If lastR < RR4_POS_FIRST Then Exit Sub

    ' band: columns I:K, same vertical extent as the donut / RL PnL chart
    Dim bandL As Double, bandT As Double, bandW As Double, bandH As Double
    bandL = ws.Columns(9).Left + 4        ' I
    bandW = ws.Columns(12).Left - bandL - 4     ' up to (not incl.) L
    bandT = ws.Rows(RR4_CHART_TOP).Top + 4
    bandH = ws.Rows(RR4_CHART_TOP + RR4_CHART_ROWS).Top - bandT - 4
    If bandH < 60 Then bandH = 60

    Dim titleH As Double: titleH = 16
    Dim lbl As Shape
    Set lbl = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, bandL, bandT, bandW, titleH)
    lbl.Name = RR4_FUN_PREFIX & "TITLE"
    lbl.Fill.Visible = msoFalse
    lbl.Line.Visible = msoFalse
    lbl.Placement = xlMove
    With lbl.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Text = "BETA EXPOSURE"
        .TextRange.Font.Name = "Consolas"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoFalse
        .TextRange.Font.Fill.ForeColor.RGB = RR4_ACCENT
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    ' 2026-09-21: pixel-aligned (see PxSnap) - top, row pitch, stage
    ' height and both edges are whole screen pixels, so every stage renders
    ' the same height instead of being anti-aliased differently.
    Dim n As Long: n = lastR - RR4_POS_FIRST + 1
    Dim barAreaT As Double: barAreaT = PxSnap(bandT + titleH + 2)
    Dim barAreaH As Double: barAreaH = bandH - titleH - 2
    Dim rowH As Double: rowH = Int(barAreaH / n / 0.75) * 0.75
    If rowH < 0.75 Then rowH = 0.75
    Dim padV As Double: padV = IIf(rowH > 6, 0.75, 0)

    ' scale bar widths off the largest |W.BETA| among the visible rows
    Dim r As Long, maxAbs As Double
    For r = RR4_POS_FIRST To lastR
        Dim v0 As Double: v0 = Abs(NumOr0(ws.cells(r, RR4_LEFT + 13).Value))
        If v0 > maxAbs Then maxAbs = v0
    Next r
    If maxAbs <= 0 Then Exit Sub

    Dim i As Long, y As Double: y = barAreaT
    For r = RR4_POS_FIRST To lastR
        i = r - RR4_POS_FIRST + 1
        Dim v As Double: v = Abs(NumOr0(ws.cells(r, RR4_LEFT + 13).Value))
        Dim frac As Double: frac = v / maxAbs
        If frac < 0.12 Then frac = 0.12    ' floor so a near-zero exposure still shows a sliver
        Dim w As Double: w = bandW * frac
        Dim h As Double: h = rowH - padV * 2
        If h < 3 Then h = rowH
        Dim x As Double: x = PxSnap(bandL + (bandW - w) / 2)
        w = PxSnap(bandL + (bandW + w) / 2) - x

        Dim tk As String: tk = CellStr(ws.cells(r, RR4_LEFT + 1).Value)
        Dim shp As Shape
        Set shp = ws.Shapes.AddShape(msoShapeTrapezoid, x, y + padV, w, h)
        shp.Name = RR4_FUN_PREFIX & i
        shp.Line.ForeColor.RGB = RGB(0, 0, 0)
        shp.Line.Weight = 1
        shp.Fill.ForeColor.RGB = FunnelColor(i)   ' 2026-09-20: same hue as the donut, darkened (less bright)
        shp.Placement = xlMove
        shp.AlternativeText = tk & "  W.BETA " & Format(v, "0.000")
        ' 2026-09-21: every stage gets a label. Wide stages keep it centred
        ' inside; narrow ones (< 46pt) get a separate textbox just right of
        ' the trapezoid, clamped so it never spills past the I:K band into
        ' the realized-PnL chart.
        Dim lblTxt As String: lblTxt = ShortTicker(tk) & " " & Format(v, "0.00")
        If w >= 46 Then
            With shp.TextFrame2
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
                .WordWrap = msoFalse
                .VerticalAnchor = msoAnchorMiddle
                .TextRange.Text = lblTxt
                .TextRange.Font.Name = "Consolas"
                .TextRange.Font.Size = 7
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            End With
        Else
            Dim sideW As Double: sideW = Len(lblTxt) * 4.2 + 4   ' Consolas 7pt ~4.2pt/char
            Dim sideX As Double: sideX = x + w + 3
            If sideX + sideW > bandL + bandW Then sideX = bandL + bandW - sideW
            sideX = PxSnap(sideX)
            Dim sideH As Double: sideH = IIf(h < 10, 10, h)
            Dim side As Shape
            Set side = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, sideX, y + padV + (h - sideH) / 2, sideW, sideH)
            side.Name = RR4_FUN_PREFIX & "L" & i
            side.Fill.Visible = msoFalse
            side.Line.Visible = msoFalse
            side.Placement = xlMove
            With side.TextFrame2
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                .WordWrap = msoFalse
                .VerticalAnchor = msoAnchorMiddle
                .TextRange.Text = lblTxt
                .TextRange.Font.Name = "Consolas"
                .TextRange.Font.Size = 7
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            End With
        End If
        y = y + rowH
    Next r
End Sub

' ================================================================
'  REALIZED PNL LINE (v4.6) - under the weight donut, same width.
'  One series: HistoryLog column D (cumulative realized PnL, TWD) against
'  column A (date), rows 2..last. Points at the sheet cells, so the line
'  follows whatever LogHistory / BackfillHistory write there.
'  v4.18 (2026-09-19): narrowed from J:S to L:S so the new beta-exposure
'  funnel (DrawFunnel) has I:K to itself.
' ================================================================
Private Sub DrawRealizedChart(ws As Worksheet)
    Dim k As Long
    For k = ws.ChartObjects.count To 1 Step -1
        If ws.ChartObjects(k).Name = RR4_RLPNL_NAME Then ws.ChartObjects(k).Delete
    Next k

    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = ThisWorkbook.Sheets(SH_HIST): On Error GoTo 0
    If wsH Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = HistLastRow(wsH)
    If lastR < HistHdrRow(wsH) + 2 Then Exit Sub          ' one point is not a line

    ' chart band (v4.18): columns L:S, full band height, right of the funnel
    Dim cL As Double, cT As Double, cW As Double, cH As Double
    cL = ws.Columns(12).Left + 4
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
        ser.Values = wsH.Range(wsH.cells(HistHdrRow(wsH) + 1, HistCol(wsH, 4)), wsH.cells(lastR, HistCol(wsH, 4)))
        ser.XValues = wsH.Range(wsH.cells(HistHdrRow(wsH) + 1, HistCol(wsH, 1)), wsH.cells(lastR, HistCol(wsH, 1)))
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
' 2026-09-25: the WATCHLIST lives in the Watch worksheet (modWatch, table
' tblWatch); this block is a read-only summary of its first 7 names.
' Read BEFORE the sheet is cleared: WatchEnsure creates tblWatch on the first
' run and seeds it from the old block, which is still intact at this point.
' Returns a 2-D Variant(1..n, 1..3) = ticker / strategy / target in Watch-page
' order; Empty when there is nothing.
Private Function ReadWatchlist(ws As Worksheet) As Variant
    Call modWatch.WatchEnsure
    ReadWatchlist = modWatch.WatchTop(RR4_WL_LAST - RR4_WL_FIRST + 1)
End Function

' Title, header, the entry row, then the saved rows with live price.
Private Sub DrawWatchlist(ws As Worksheet, wl As Variant)
    ' 2026-09-21: grey how-to hint dropped (user request) - title only
    With ws.cells(RR4_WL_TITLE, RR4_LEFT + 1)
        .Value = "WATCHLIST"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 10
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
    ' closing rule under the last saved row, separating it from TO-DO
    With ws.Range(ws.cells(RR4_WL_LAST, RR4_LEFT + 1), ws.cells(RR4_WL_LAST, RR4_LEFT + 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With
End Sub

' 2026-09-25: the old entry row (RR4_WL_ENTRY) is retired - names are typed
' on the Watch page. Keep the row blank and black so no input styling is left.
Private Sub WatchlistPaintEntryRow(ws As Worksheet)
    With ws.Range(ws.cells(RR4_WL_ENTRY, RR4_LEFT + 1), ws.cells(RR4_WL_ENTRY, RR4_LEFT + 4))
        .ClearContents
        .Interior.Color = RGB(0, 0, 0)
        .Borders(xlEdgeRight).LineStyle = xlNone
    End With
    Dim c As Long
    For c = 1 To 3
        ws.cells(RR4_WL_ENTRY, RR4_LEFT + c).Borders(xlEdgeRight).LineStyle = xlNone
    Next c
End Sub

' Black vertical rules between adjacent grey input cells of an entry row
' (B..B+n-1), so each field reads as its own box (2026-09-21).
Private Sub SeparateInputCells(ws As Worksheet, ByVal r As Long, ByVal n As Long)
    Dim c As Long
    For c = 1 To n - 1
        With ws.cells(r, RR4_LEFT + c).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .Weight = xlMedium
        End With
    Next c
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
    If px > 0 Then Call modWatch.WatchCacheLast(tk, px)     ' keep tblWatch's LAST cache current
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

' ================================================================
'  TO-DO (B36:E39, 2026-09-21) - see the RR4_TD_* constants
' ================================================================
' Saved rows, read BEFORE the sheet is cleared; only trusted when the
' title is in place. Returns Variant(1..n, 1..3) = ticker / task / due
' (due = Date or Empty), or Empty when there is nothing.
Private Function ReadTodo(ws As Worksheet) As Variant
    If UCase(CellStr(ws.cells(RR4_TD_TITLE, RR4_LEFT + 1).Value)) <> "TO-DO" Then Exit Function
    Dim tmp(1 To 4, 1 To 3) As Variant, n As Long, r As Long
    For r = RR4_TD_FIRST To RR4_TD_LAST
        Dim tsk As String: tsk = CellStr(ws.cells(r, RR4_LEFT + 2).Value)
        If tsk <> "" Then
            n = n + 1
            tmp(n, 1) = UCase(CellStr(ws.cells(r, RR4_LEFT + 1).Value))
            tmp(n, 2) = tsk
            tmp(n, 3) = TodoDue(ws.cells(r, RR4_LEFT + 3).Value)
        End If
    Next r
    If n = 0 Then Exit Function
    Dim out() As Variant: ReDim out(1 To n, 1 To 3)
    For r = 1 To n
        out(r, 1) = tmp(r, 1): out(r, 2) = tmp(r, 2): out(r, 3) = tmp(r, 3)
    Next r
    ReadTodo = out
End Function

' A due-date cell value as a Date, or Empty when it is not a date.
Private Function TodoDue(ByVal v As Variant) As Variant
    TodoDue = Empty
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If IsDate(v) Then TodoDue = CDate(Int(CDbl(CDate(v)))): Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 30000 And CDbl(v) < 80000 Then TodoDue = CDate(Int(CDbl(v)))
    End If
End Function

' Title + column names, the entry row, then the saved rows (sorted).
Private Sub DrawTodo(ws As Worksheet, td As Variant)
    With ws.cells(RR4_TD_TITLE, RR4_LEFT + 1)
        .Value = "TO-DO"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
    End With
    Dim hdr As Variant: hdr = Array("TASK", "DUE", "DTE")
    Dim c As Long
    For c = 0 To 2
        With ws.cells(RR4_TD_TITLE, RR4_LEFT + 2 + c)
            .Value = hdr(c)
            .Font.Color = RGB(0, 200, 255)
            .Font.Size = 9
            .Font.Bold = False
            .HorizontalAlignment = IIf(c = 0, xlLeft, xlCenter)
        End With
    Next c
    Call TodoPaintEntryRow(ws)
    Call TodoWriteList(ws, td)
End Sub

' The entry row: B ticker / C task / D due, three dark input cells.
Private Sub TodoPaintEntryRow(ws As Worksheet)
    Dim c As Long
    For c = 1 To 3
        With ws.cells(RR4_TD_ENTRY, RR4_LEFT + c)
            .Value = ""
            .Interior.Color = RR4_INPUT_BG
            .Font.Color = RR4_INPUT_FG
            .Font.Bold = (c = 1)
            .HorizontalAlignment = IIf(c = 2, xlLeft, xlCenter)
            .NumberFormat = IIf(c = 3, "yyyy/m/d", "@")
        End With
    Next c
    With ws.cells(RR4_TD_ENTRY, RR4_LEFT + 4)
        .Value = ""
        .Interior.Color = RGB(0, 0, 0)
    End With
    Call SeparateInputCells(ws, RR4_TD_ENTRY, 3)
End Sub

' Sort by due date (undated last, ties keep order) and write rows 38-39;
' DTE = due - today, orange once negative.
Private Sub TodoWriteList(ws As Worksheet, td As Variant)
    Dim n As Long: If IsArray(td) Then n = UBound(td, 1)
    Dim ord() As Long, i As Long, j As Long, t As Long
    If n > 0 Then
        ReDim ord(1 To n)
        For i = 1 To n: ord(i) = i: Next i
        For i = 2 To n
            t = ord(i): j = i - 1
            Do While j >= 1
                If Not TodoBefore(td, t, ord(j)) Then Exit Do
                ord(j + 1) = ord(j): j = j - 1
            Loop
            ord(j + 1) = t
        Next i
    End If

    Dim r As Long, k As Long
    For r = RR4_TD_FIRST To RR4_TD_LAST
        k = r - RR4_TD_FIRST + 1
        Dim row4 As Range
        Set row4 = ws.Range(ws.cells(r, RR4_LEFT + 1), ws.cells(r, RR4_LEFT + 4))
        row4.ClearContents
        row4.Interior.Color = RGB(0, 0, 0)
        row4.Font.Color = RGB(221, 221, 221)
        row4.Font.Bold = False
        ws.cells(r, RR4_LEFT + 1).NumberFormat = "@"
        ws.cells(r, RR4_LEFT + 2).NumberFormat = "@"
        ws.cells(r, RR4_LEFT + 3).NumberFormat = "yyyy/m/d"
        ws.cells(r, RR4_LEFT + 4).NumberFormat = "0"
        ws.cells(r, RR4_LEFT + 1).HorizontalAlignment = xlCenter
        ws.cells(r, RR4_LEFT + 2).HorizontalAlignment = xlLeft
        ws.cells(r, RR4_LEFT + 3).HorizontalAlignment = xlCenter
        ws.cells(r, RR4_LEFT + 4).HorizontalAlignment = xlCenter
        If k <= n Then
            i = ord(k)
            ws.cells(r, RR4_LEFT + 1).Value = td(i, 1)
            ws.cells(r, RR4_LEFT + 1).Font.Color = RR4_ACCENT
            ws.cells(r, RR4_LEFT + 1).Font.Bold = True
            ws.cells(r, RR4_LEFT + 2).Value = td(i, 2)
            If Not IsEmpty(td(i, 3)) Then
                ws.cells(r, RR4_LEFT + 3).Value = td(i, 3)
                Dim dte As Long: dte = CLng(CDbl(td(i, 3)) - CDbl(Date))
                ws.cells(r, RR4_LEFT + 4).Value = dte
                If dte < 0 Then
                    ws.cells(r, RR4_LEFT + 4).Font.Color = RR4_ACCENT
                    ws.cells(r, RR4_LEFT + 4).Font.Bold = True
                End If
            End If
        End If
    Next r
End Sub

' True when item a sorts before item b (earlier due; dated before undated).
Private Function TodoBefore(td As Variant, ByVal a As Long, ByVal b As Long) As Boolean
    If IsEmpty(td(a, 3)) Then Exit Function
    If IsEmpty(td(b, 3)) Then TodoBefore = True: Exit Function
    TodoBefore = (CDbl(td(a, 3)) < CDbl(td(b, 3)))
End Function

' Called by the sheet code when B/C/D of the TO-DO entry row changes.
' Commits once the task (C) is in and the row is finished: Enter in the
' due cell D (a date, or anything non-date such as "-" for no date), or
' Enter in C when D already holds a date.
Public Sub TodoCommitEntry(ws As Worksheet, ByVal changedCol As Long)
    Dim tk As String: tk = UCase(Trim(CellStr(ws.cells(RR4_TD_ENTRY, RR4_LEFT + 1).Value)))
    Dim tsk As String: tsk = Trim(CellStr(ws.cells(RR4_TD_ENTRY, RR4_LEFT + 2).Value))
    Dim due As Variant: due = TodoDue(ws.cells(RR4_TD_ENTRY, RR4_LEFT + 3).Value)
    If tsk = "" Then Exit Sub
    If changedCol = RR4_LEFT + 1 Then Exit Sub
    If changedCol = RR4_LEFT + 2 And IsEmpty(due) Then Exit Sub

    Dim td As Variant: td = ReadTodo(ws)
    Dim n As Long: If IsArray(td) Then n = UBound(td, 1)
    If n >= RR4_TD_LAST - RR4_TD_FIRST + 1 Then
        Call NavNotify("TO-DO full (" & (RR4_TD_LAST - RR4_TD_FIRST + 1) & " rows) - double-click a row to mark it done", True)
        Exit Sub
    End If
    Dim nw() As Variant: ReDim nw(1 To n + 1, 1 To 3)
    Dim i As Long
    For i = 1 To n
        nw(i, 1) = td(i, 1): nw(i, 2) = td(i, 2): nw(i, 3) = td(i, 3)
    Next i
    nw(n + 1, 1) = tk: nw(n + 1, 2) = tsk: nw(n + 1, 3) = due
    Call TodoWriteList(ws, nw)
    Call TodoPaintEntryRow(ws)
    ws.cells(RR4_TD_ENTRY, RR4_LEFT + 1).Select
    Call NavNotify("TO-DO + " & tsk)
End Sub

' Called by the sheet code on a double-click inside the saved rows: the
' task is done - drop it and close the gap.
Public Sub TodoDeleteRow(ws As Worksheet, ByVal r As Long)
    If r < RR4_TD_FIRST Or r > RR4_TD_LAST Then Exit Sub
    Dim tsk As String: tsk = CellStr(ws.cells(r, RR4_LEFT + 2).Value)
    If tsk = "" Then Exit Sub
    Dim td As Variant: td = ReadTodo(ws)
    Dim n As Long: If IsArray(td) Then n = UBound(td, 1)
    Dim skip As Long: skip = r - RR4_TD_FIRST + 1
    Dim nw As Variant, m As Long, i As Long
    If n > 1 Then
        Dim tmp() As Variant: ReDim tmp(1 To n - 1, 1 To 3)
        For i = 1 To n
            If i <> skip Then
                m = m + 1
                tmp(m, 1) = td(i, 1): tmp(m, 2) = td(i, 2): tmp(m, 3) = td(i, 3)
            End If
        Next i
        nw = tmp
    End If
    Call TodoWriteList(ws, nw)
    Call NavNotify("TO-DO done - " & tsk)
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

' Darkened version of a colour, each channel scaled by factor (0..1).
Private Function DarkenColor(ByVal clr As Long, ByVal factor As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = clr Mod 256
    g = (clr \ 256) Mod 256
    b = (clr \ 65536) Mod 256
    DarkenColor = RGB(r * factor, g * factor, b * factor)
End Function

' Funnel stage n's colour: same hue as the WEIGHT donut's slice n, darkened
' (2026-09-20, user) - the donut's saturated palette looked too bright on
' the taller funnel bars.
Private Function FunnelColor(ByVal n As Long) As Long
    FunnelColor = DarkenColor(DonutColor(n), 0.6)
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

' NET EXPOSURE from the summary block (the page's total market value, TWD)
Public Function RR4NetExposure() As Double
    On Error Resume Next
    RR4NetExposure = CDbl(ThisWorkbook.Sheets(SH_PORT).cells(RR4_TOP + 16, RR4_LEFT + 6).Value)
    On Error GoTo 0
End Function

Public Function RR4PositionCount() As Long
    Dim lastR As Long: lastR = LastPositionRow(ThisWorkbook.Sheets(SH_PORT))
    If lastR >= RR4_POS_FIRST Then RR4PositionCount = lastR - RR4_POS_FIRST + 1
End Function

' ================================================================
'  Small helpers (v4)
' ================================================================
' UPSIDE and DOWNSIDE are typed by hand, so they have to survive the
' sheet clear. Found by header text (whatever row the log is on), keyed
' by ticker so an ARRANGE between two UPs does not mis-assign them.
Private Function ReadHandColumn(ws As Worksheet, ByVal headerText As String) As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    Dim hdrRow As Long, hdrCol As Long, r As Long, c As Long
    For r = 1 To 60
        For c = 1 To 20
            If UCase(CellStr(ws.cells(r, c).Value)) = UCase(headerText) Then
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
    Set ReadHandColumn = m
End Function

' HistoryLog column C (cumulative PnL) on the last row dated before today
Private Function PrevDayCumPnL() As Variant
    PrevDayCumPnL = Empty
    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = ThisWorkbook.Sheets(SH_HIST): On Error GoTo 0
    If wsH Is Nothing Then Exit Function
    Dim r As Long
    For r = HistLastRow(wsH) To HistHdrRow(wsH) + 1 Step -1
        Dim dv As Variant: dv = wsH.cells(r, HistCol(wsH, 1)).Value
        If IsDate(dv) Then
            If Int(CDate(dv)) < Date Then
                Dim cv As Variant: cv = wsH.cells(r, HistCol(wsH, 3)).Value
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
    lastR = RealLastRow(wsR)
    For r = RealHdrRow(wsR) + 1 To lastR
        Dim dv As Variant: dv = wsR.cells(r, RealCol(wsR, 9)).Value
        If IsDate(dv) Then
            If Int(CDate(dv)) < d Then RealizedBefore = RealizedBefore + NumOr0(wsR.cells(r, RealCol(wsR, 8)).Value)
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

' A hand-typed cell read back as text: a number when it parses (so the
' UPSIDE / DOWNSIDE formats and the P.TARGET / SWING RISK formulas see it),
' otherwise the text as typed.
Private Function HandNumber(ByVal s As String) As Variant
    If IsNumeric(s) Then HandNumber = CDbl(s) Else HandNumber = s
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

' ================================================================
'  LIVE AUTO-REFRESH (cell-only, 2026-09-23) - called every TICK_SEC by
'  modLiveRefresh.Tick once the user has sat on this sheet for a minute.
'  Deliberately does NOT call BuildPositions/CalcPositions or redraw
'  totals/donut/weight-bars/TOP EXPOSURE - RebuildPortfolioDashboard
'  measured ~22s per run, far too slow to fire every 15s. Only the LAST
'  column (position table) and the price column (watchlist) get new
'  numbers, only for tickers whose market is open right now; % CHG /
'  UNRL PNL / totals are left exactly as the last UP computed them.
' ================================================================
Public Sub RefreshLivePricesLite()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_PORT)

    Dim twOpen As Boolean: twOpen = modMarketHours.IsTWMarketOpen()
    Dim usOpen As Boolean: usOpen = modMarketHours.IsUSMarketOpen()
    If Not twOpen And Not usOpen Then
        Call ShowLiveRefreshStatus(ws, queried:=False, gotPrice:=False)
        Exit Sub
    End If

    Dim queriedCount As Long: queriedCount = 0

    ' one batched MIS request covers every TW ticker on screen, same
    ' pattern as RebuildPortfolioDashboard's own prefetch
    If twOpen Then
        Dim twTickers() As String: twTickers = CollectOpenMarketTWTickers(ws)
        modMISPrice.PrefetchMISPrices twTickers
    End If

    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Dim prevScr    As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fin

    ' Ticker-board flash (2026-09-23; switched from a border outline, then
    ' briefly to a full-cell background flash, then - per explicit user
    ' feedback ("底色不用改變都是黑色，字變動就好") - to a FONT-colour-only
    ' flash. Background/Interior.Color is never touched by the tick-flash
    ' at all (it stays exactly as already painted - black row stripes for
    ' the position table, the P.TARGET hit/no-hit fill for the watchlist,
    ' set unconditionally below regardless of whether this tick flashes).
    ' Every cell whose price actually changed this tick gets its Font.Color
    ' set to the up/down colour INSTANTLY (no tween animation - user wants
    ' "瞬間跳變、乾脆俐落"), all of them together; once every changed cell
    ' is written ScreenUpdating flips back on and the flash holds briefly,
    ' then every cell's Font.Color is restored in one more pass - to its
    ' own normal resting colour (LAST: light grey; %CHG/UNRL PNL: the
    ' muted PnL colour; watchlist LAST: accent-orange-bold if P.TARGET hit,
    ' else light grey - captured/computed per cell, never assumed uniform,
    ' since LAST/%CHG/UNRL PNL each rest at a different colour). Cells with
    ' no prior price (first-ever populate) or an unchanged price are left
    ' alone - flashing only marks a real tick, like the board this is
    ' modelled on.
    Dim flashCell() As Range, flashRestoreColor() As Long, flashCount As Long
    ' each changed position row can now flash up to 3 cells (LAST/%CHG/UNRL
    ' PNL, 2026-09-23), each changed watchlist row up to 1 - size generously
    Dim maxFlash As Long: maxFlash = (RR4_WL_LAST - RR4_WL_FIRST + 1) + 1500
    ReDim flashCell(1 To maxFlash)
    ReDim flashRestoreColor(1 To maxFlash)
    flashCount = 0

    ' 2026-09-23: sum of every position's UNRL PNL change this tick (TWD
    ' terms, same conversion already applied to each row's own cell below)
    ' - lets the SUMMARY block's UNREALISED PNL / NET EXPOSURE / UNREALISED
    ' PNL % move on every 15s tick too, without needing a full UP (which
    ' this "lite" refresh deliberately skips - see module header). Cost
    ' basis does not change on a price tick, so market value and
    ' unrealised PnL move by exactly the same delta.
    Dim totalUnrlDeltaTWD As Double: totalUnrlDeltaTWD = 0

    ' ---- position table: RR4_POS_HDR+1 downward while TICKER is non-blank ----
    Dim r As Long: r = RR4_POS_HDR + 1
    Do While Len(Trim(CStr(ws.cells(r, RR4_LEFT + 1).Value))) > 0
        Dim rawTk As String: rawTk = CStr(ws.cells(r, RR4_LEFT + 1).Value)
        Dim normTw As String: normTw = NormalizeTWTicker(rawTk)
        Dim isTw As Boolean: isTw = (normTw <> "")
        If (isTw And twOpen) Or (Not isTw And usOpen) Then
            Dim queryTk As String: queryTk = IIf(isTw, normTw, UCase(Trim(rawTk)))
            Dim px As Double: px = modLivePrice.GetLivePrice(queryTk)
            If px > 0 Then
                queriedCount = queriedCount + 1
                Dim lastCell As Range: Set lastCell = ws.cells(r, RR4_LEFT + 9)
                Dim oldPx As Double: oldPx = NumOr0(lastCell.Value)
                lastCell.Value = px
                If oldPx > 0 And px <> oldPx Then
                    Dim tickFlashColor As Long: tickFlashColor = PnLColor(px - oldPx)

                    ' LAST always rests at the same light grey regardless of
                    ' direction (set once, here, by WriteOnePositionRow) -
                    ' flash it, then restore to that.
                    flashCount = flashCount + 1
                    Set flashCell(flashCount) = lastCell
                    flashRestoreColor(flashCount) = RGB(221, 221, 221)
                    lastCell.Font.Color = tickFlashColor

                    ' 2026-09-23: % CHG / UNRL PNL recomputed from the ENTRY
                    ' PX / SHARES / C6 already sitting on the sheet (no
                    ' BuildPositions/CalcPositions re-run - that's the whole
                    ' point of "lite") and flashed in the same beat as LAST.
                    ' Each of the 3 cells is flashed/restored individually
                    ' (not as one merged range) because they rest at
                    ' different colours - LAST is always grey, %CHG/UNRL
                    ' PNL rest at the muted PnL colour.
                    Dim entryPx2 As Double: entryPx2 = NumOr0(ws.cells(r, RR4_LEFT + 8).Value)
                    If entryPx2 > 0 Then
                        Dim shares2 As Double: shares2 = NumOr0(ws.cells(r, RR4_LEFT + 7).Value)
                        Dim newChgPct As Double: newChgPct = (px - entryPx2) / entryPx2
                        Dim newUnrl As Double: newUnrl = (px - entryPx2) * shares2
                        If GetCurrencyType(queryTk) = "USD" Then
                            Dim exRateNow As Double: exRateNow = NumOr0(ws.Range(RR4_FX_CELL).Value)
                            If exRateNow > 0 Then newUnrl = newUnrl * exRateNow
                        End If
                        Dim pnlFontColor As Long: pnlFontColor = PnLColorMuted(newUnrl)

                        Dim chgPctCell As Range: Set chgPctCell = ws.cells(r, RR4_LEFT + 10)
                        chgPctCell.Value = newChgPct
                        flashCount = flashCount + 1
                        Set flashCell(flashCount) = chgPctCell
                        flashRestoreColor(flashCount) = pnlFontColor
                        chgPctCell.Font.Color = tickFlashColor

                        Dim unrlCell As Range: Set unrlCell = ws.cells(r, RR4_LEFT + 11)
                        Dim oldUnrlCellVal As Double: oldUnrlCellVal = NumOr0(unrlCell.Value)
                        unrlCell.Value = newUnrl
                        totalUnrlDeltaTWD = totalUnrlDeltaTWD + (newUnrl - oldUnrlCellVal)
                        flashCount = flashCount + 1
                        Set flashCell(flashCount) = unrlCell
                        flashRestoreColor(flashCount) = pnlFontColor
                        unrlCell.Font.Color = tickFlashColor
                    End If
                End If
            End If
        End If
        r = r + 1
    Loop

    ' ---- watchlist: RR4_WL_FIRST..RR4_WL_LAST, price column RR4_LEFT+4 ----
    ' the P.TARGET-hit highlight (RefreshWatchlistRow's own orange/black
    ' rule) can flip with the new price, so the resting colour here is
    ' RECOMPUTED from the new price, never just "whatever was there before".
    Dim wr As Long
    For wr = RR4_WL_FIRST To RR4_WL_LAST
        Dim wRaw As String: wRaw = CStr(ws.cells(wr, RR4_LEFT + 1).Value)
        If Len(Trim(wRaw)) > 0 Then
            Dim wNormTw As String: wNormTw = NormalizeTWTicker(wRaw)
            Dim wIsTw As Boolean: wIsTw = (wNormTw <> "")
            If (wIsTw And twOpen) Or (Not wIsTw And usOpen) Then
                Dim wQueryTk As String: wQueryTk = IIf(wIsTw, wNormTw, UCase(Trim(wRaw)))
                Dim wpx As Double: wpx = modLivePrice.GetLivePrice(wQueryTk)
                If wpx > 0 Then
                    queriedCount = queriedCount + 1
                    Dim priceCell As Range: Set priceCell = ws.cells(wr, RR4_LEFT + 4)
                    Dim oldWpx As Double: oldWpx = NumOr0(priceCell.Value)
                    priceCell.Value = wpx
                    Dim tgt As Double: tgt = NumOr0(ws.cells(wr, RR4_LEFT + 3).Value)
                    Dim hit As Boolean: hit = (wpx > 0 And tgt > 0 And wpx <= tgt)
                    ' the P.TARGET-hit fill (orange/black background, and its
                    ' matching resting font colour) is independent of the
                    ' tick-flash - always set to what the new price says it
                    ' should be, whether or not this tick also flashes.
                    ' Interior.Color is NEVER touched by the flash itself.
                    priceCell.Interior.Color = IIf(hit, RGB(60, 30, 0), RGB(0, 0, 0))
                    Dim restFontColor As Long: restFontColor = IIf(hit, RR4_ACCENT, RGB(221, 221, 221))
                    If oldWpx > 0 And wpx <> oldWpx Then
                        flashCount = flashCount + 1
                        Set flashCell(flashCount) = priceCell
                        flashRestoreColor(flashCount) = restFontColor
                        priceCell.Font.Color = PnLColor(wpx - oldWpx)
                    Else
                        priceCell.Font.Color = restFontColor
                    End If
                End If
            End If
        End If
    Next wr

    ' 2026-09-23: push the accumulated UNRL PNL delta into the SUMMARY
    ' block (UNREALISED PNL, NET EXPOSURE, UNREALISED PNL %) - see the
    ' totalUnrlDeltaTWD comment above. Cost basis (TWD) is derived from
    ' the summary's own pre-tick UNREALISED PNL / NET EXPOSURE cells
    ' (cost = market value - unrealised PnL) rather than recomputed from
    ' the position table, since CalcPositions is exactly the expensive
    ' full-rebuild path this "lite" refresh exists to avoid.
    If totalUnrlDeltaTWD <> 0 Then
        Dim unrlPnlCell As Range: Set unrlPnlCell = ws.cells(RR4_TOP + 19, RR4_LEFT + 2)
        Dim netExpCell As Range: Set netExpCell = ws.cells(RR4_TOP + 16, RR4_LEFT + 6)
        Dim unrlPctCell As Range: Set unrlPctCell = ws.cells(RR4_TOP + 19, RR4_LEFT + 6)

        Dim oldTotalUnrlS As Double: oldTotalUnrlS = NumOr0(unrlPnlCell.Value)
        Dim oldTotalMktS As Double: oldTotalMktS = NumOr0(netExpCell.Value)
        Dim totalCostS As Double: totalCostS = oldTotalMktS - oldTotalUnrlS

        Dim newTotalUnrlS As Double: newTotalUnrlS = oldTotalUnrlS + totalUnrlDeltaTWD
        unrlPnlCell.Value = newTotalUnrlS
        unrlPnlCell.Font.Color = GainLossColor(newTotalUnrlS)

        netExpCell.Value = oldTotalMktS + totalUnrlDeltaTWD

        If totalCostS > 0 Then
            Dim newUnrlPctS As Double: newUnrlPctS = newTotalUnrlS / totalCostS
            unrlPctCell.Value = newUnrlPctS
            unrlPctCell.Font.Color = GainLossColor(newUnrlPctS)
        End If
    End If

    If flashCount > 0 Then
        Application.ScreenUpdating = True
        ' 2026-09-23: every changed cell's TEXT already switched to its
        ' flash colour above, all in the same pass - so the "whole board"
        ' pops together the instant ScreenUpdating flips back on. 300ms
        ' hold is deliberately short ("瞬間跳變、乾脆俐落" - user explicitly
        ' did not want a lingering effect). Background is never touched.
        Call PauseFlashMS(300)
        Dim k As Long
        For k = 1 To flashCount
            flashCell(k).Font.Color = flashRestoreColor(k)
        Next k
    End If

    Call ShowLiveRefreshStatus(ws, queried:=True, gotPrice:=(queriedCount > 0))

Fin:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScr
End Sub

Private Sub PauseFlashMS(ByVal ms As Long)
    Dim t As Single: t = Timer + ms / 1000
    Do While Timer < t: DoEvents: Loop
End Sub

' ================================================================
'  NAV-BAR HEARTBEAT (2026-09-23) - the 15s auto-refresh was otherwise
'  completely silent, so this writes a small status word into the RR4
'  nav bar itself (row 2, column P - clear of the C2 command cell, the
'  code-line overflow in rows 3/4, and the T2/T3 config cells at column
'  20), then clears it after a couple of seconds. Deliberately separate
'  from modNav.NavNotify (the Excel status-bar line UP/V!/etc. already
'  use) - this is a cell people can actually see without looking at the
'  bottom of the window, and it only ever reflects the auto-refresh tick.
' ================================================================
' 2026-09-23: this used to write a status word then blank itself back to ""
' after ~1.5s. Per user feedback ("update 的標示都存在") it now stays on
' screen permanently - the last status is always visible, never cleared -
' and rests at a pale red rather than the old green/grey. Only on an
' actual successful update does it briefly brighten to a fuller red, then
' settle back to the pale resting shade; MARKET CLOSED / FETCH FAILED just
' sit at the pale colour with no flash (nothing genuinely updated).
Private Sub ShowLiveRefreshStatus(ws As Worksheet, ByVal queried As Boolean, ByVal gotPrice As Boolean)
    Dim cell As Range: Set cell = ws.cells(LIVE_STATUS_ROW, LIVE_STATUS_COL)
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Dim restColor As Long: restColor = RGB(204, 102, 102)   ' pale red, always-on resting colour
    Dim flashColor As Long: flashColor = RGB(255, 40, 40)   ' brighter red, brief pop on a real update

    cell.NumberFormat = "@"
    cell.Font.Name = "Consolas"
    cell.Font.Size = 9
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlLeft
    cell.Interior.Color = RGB(0, 0, 0)

    If Not queried Then
        cell.Value = "MARKET CLOSED"
        cell.Font.Color = restColor
    ElseIf gotPrice Then
        cell.Value = ChrW(&H2713) & " UPDATED " & Format(Now, "hh:mm:ss")
        cell.Font.Color = flashColor
        Application.ScreenUpdating = True
        Call PauseFlashMS(300)
        cell.Font.Color = restColor
    Else
        cell.Value = "PRICE FETCH FAILED " & Format(Now, "hh:mm:ss")
        cell.Font.Color = restColor
    End If

    Application.ScreenUpdating = prevScr
End Sub

' Every TW ticker currently on screen (position table + watchlist),
' normalized to include .TW/.TWO - for a single batched MIS prefetch.
Private Function CollectOpenMarketTWTickers(ws As Worksheet) As String()
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

    Dim r As Long: r = RR4_POS_HDR + 1
    Do While Len(Trim(CStr(ws.cells(r, RR4_LEFT + 1).Value))) > 0
        Dim tk As String: tk = NormalizeTWTicker(CStr(ws.cells(r, RR4_LEFT + 1).Value))
        If tk <> "" And Not seen.Exists(tk) Then seen(tk) = 1
        r = r + 1
    Loop

    Dim wr As Long
    For wr = RR4_WL_FIRST To RR4_WL_LAST
        Dim wtk As String: wtk = NormalizeTWTicker(CStr(ws.cells(wr, RR4_LEFT + 1).Value))
        If wtk <> "" And Not seen.Exists(wtk) Then seen(wtk) = 1
    Next wr

    Dim out() As String
    If seen.Count = 0 Then
        ReDim out(0 To -1)
    Else
        ReDim out(0 To seen.Count - 1)
        Dim i As Long: i = 0
        Dim k As Variant
        For Each k In seen.keys
            out(i) = CStr(k): i = i + 1
        Next k
    End If
    CollectOpenMarketTWTickers = out
End Function

' 2026-09-23: 收集這一輪 UP 需要報價的全部台股代號（持倉 + watchlist +
' ticker panel），交給 modMISPrice.PrefetchMISPrices 一次批次查詢。正規化
' 規則跟 Attach.GetStockPrice 的 STEP 1 一致（裸數字代號補 .TW），這樣這裡
' 收集到的 key 才會跟 GetStockPrice 之後實際查快取用的 key 一致。
Private Function CollectTWTickersForPrefetch(positions As Object, wl As Variant, tiTicker As String) As String()
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

    Dim tkv As Variant
    For Each tkv In positions.keys
        Dim pp As Long: pp = InStr(CStr(tkv), "|")
        Dim t As String: t = IIf(pp > 0, Left(CStr(tkv), pp - 1), CStr(tkv))
        t = NormalizeTWTicker(t)
        If t <> "" And Not seen.Exists(t) Then seen(t) = 1
    Next tkv

    If Not IsEmpty(wl) Then
        Dim r As Long
        For r = LBound(wl, 1) To UBound(wl, 1)
            Dim wt As String: wt = NormalizeTWTicker(CStr(wl(r, 1)))
            If wt <> "" And Not seen.Exists(wt) Then seen(wt) = 1
        Next r
    End If

    Dim tt As String: tt = NormalizeTWTicker(tiTicker)
    If tt <> "" And Not seen.Exists(tt) Then seen(tt) = 1

    Dim out() As String
    If seen.Count = 0 Then
        ReDim out(0 To -1)
    Else
        ReDim out(0 To seen.Count - 1)
        Dim i As Long: i = 0
        Dim k As Variant
        For Each k In seen.keys
            out(i) = CStr(k): i = i + 1
        Next k
    End If
    CollectTWTickersForPrefetch = out
End Function

' 台股才回傳正規化後的代號（含 .TW/.TWO），美股/期貨代號一律回空字串跳過。
Private Function NormalizeTWTicker(raw As String) As String
    Dim u As String: u = UCase(Trim(raw))
    If u = "" Then NormalizeTWTicker = "": Exit Function
    If InStr(u, ".TW") = 0 And InStr(u, ".TWO") = 0 Then
        If IsNumeric(u) Then
            u = u & ".TW"
        Else
            NormalizeTWTicker = ""
            Exit Function
        End If
    End If
    NormalizeTWTicker = u
End Function

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
    Set m_sellInfo = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = TrLastRow(wsTr)
    If lastRow <= TrHdrRow(wsTr) Then Set BuildPositions = dict: Exit Function

    Dim data As Variant                         ' page columns B..N, wherever the bar puts them
    data = wsTr.Range(wsTr.cells(TrHdrRow(wsTr) + 1, TrCol(wsTr, 2)), wsTr.cells(lastRow, TrCol(wsTr, 14))).Value

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim Ticker As String: Ticker = CStr(data(i, 2))
        Dim Action As String: Action = CStr(data(i, 3))
        Dim shares As Double: shares = val(data(i, 4))
        Dim price  As Double: price = val(data(i, 5))
        Dim fee    As Double: fee = val(data(i, 6))
        Dim tax    As Double: tax = val(data(i, 7))
        Dim netAmt As Double: netAmt = val(data(i, 8))
        Dim Sector As String: Sector = CStr(data(i, 9))
        Dim strat  As String: strat = CStr(data(i, 11))
        Dim Beta   As Double: Beta = val(data(i, 12))
        Dim broker As String
        On Error Resume Next
        broker = Trim(CStr(data(i, 13)))
        On Error GoTo 0
        If broker = "" Then broker = "Default"

        If Ticker <> "" Then
            ' The key keeps the ticker EXACTLY as logged, suffix and all.
            ' Stripping .TW/.TWO here (tried 2026-09-18) looks harmless but
            ' breaks the whole page: CalcPositions parses this key back out
            ' and hands it to GetCurrencyType, which decides TWD vs USD
            ' purely on the .TW suffix - so every TW position turned "USD"
            ' and its market value got multiplied by USD/TWD, blowing up
            ' NET EXPOSURE and every WT%. Same trap as TickerInsight v2.4.
            Dim posKey As String: posKey = Ticker & "|" & broker
            If Not dict.Exists(posKey) Then
                dict.Add posKey, Array(0#, 0#, 0#, 0#, 0#, 0#, "", 0#, "", 0#, broker, New Collection)
            End If
            Dim d As Variant: d = dict(posKey)
            If strat <> "" Then d(8) = strat
            If Beta <> 0 Then d(9) = Beta
            If Sector <> "" Then d(6) = Sector
            d(10) = broker

            Dim lots As Collection: Set lots = d(11)
            Dim tDate As Variant: tDate = data(i, 1)

            Select Case UCase(Action)
                Case "BUY"
                    ' 2026-09-18 (owner's decision): the cost basis is the
                    ' ACQUISITION cost - shares*price plus the BUY-side fee
                    ' and tax. It deliberately no longer reads Net_Amount,
                    ' which frmTransaction inflates with an ESTIMATE of the
                    ' future SELL-side fee and tax. That estimate pushed
                    ' ENTRY PX above every fill price (3653 by 18.7/share,
                    ' 7610 by 4.4) and understated UNRL PNL by the same
                    ' amount. Older rows also predate the estimate, so the
                    ' column is not even one convention across the log.
                    Dim buyCost As Double: buyCost = shares * price + fee + tax
                    ' hand-typed row with no Price: fall back to Net_Amount
                    ' rather than costing the lot at (almost) zero
                    If shares * price <= 0 Then buyCost = netAmt
                    Dim cps As Double: If shares > 0 Then cps = buyCost / shares Else cps = 0
                    Dim lotDateSer As Long: lotDateSer = IIf(IsDate(tDate), CLng(CDate(tDate)), 0)
                    lots.Add Array(lotDateSer, shares, cps)

                Case "SELL"
                    Dim remaining As Double: remaining = shares
                    Dim exitSer As Long: exitSer = IIf(IsDate(tDate), CLng(CDate(tDate)), 0)
                    Dim firstIn As Long, lastIn As Long, shDays As Double, usedSh As Double, usedCost As Double
                    firstIn = 0: lastIn = 0: shDays = 0: usedSh = 0: usedCost = 0
                    Do While remaining > 0.0001 And lots.count > 0
                        Dim frontLot As Variant: frontLot = lots(1)
                        Dim lotSh As Double: lotSh = frontLot(1)
                        Dim lotCps As Double: lotCps = frontLot(2)
                        Dim took As Double
                        If lotSh <= remaining + 0.0001 Then took = lotSh Else took = remaining
                        ' what this sell consumed, for the DAILY LOG line
                        If frontLot(0) > 0 Then
                            If firstIn = 0 Or frontLot(0) < firstIn Then firstIn = frontLot(0)
                            If frontLot(0) > lastIn Then lastIn = frontLot(0)
                            If exitSer > 0 Then shDays = shDays + took * (exitSer - frontLot(0))
                        End If
                        usedSh = usedSh + took
                        usedCost = usedCost + took * lotCps
                        If lotSh <= remaining + 0.0001 Then
                            remaining = remaining - lotSh
                            lots.Remove 1
                        Else
                            lots.Add Array(frontLot(0), lotSh - remaining, lotCps), Before:=1
                            lots.Remove 2
                            remaining = 0
                        End If
                    Loop
                    Dim avgDays As Double: If usedSh > 0 Then avgDays = shDays / usedSh
                    m_sellInfo(TrHdrRow(wsTr) + i) = Array(firstIn, lastIn, avgDays, usedCost)

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

' v4.14: USD/TWD spot from the Yahoo chart API (TWD=X = 1 USD in TWD).
' meta.regularMarketPrice is the latest quote; 0 when anything fails, so
' the caller falls back to the typed C6 value.
Private Function FetchLiveFx() As Double
    Dim http As Object, resp As String
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://query1.finance.yahoo.com/v8/finance/chart/TWD=X?interval=1d&range=5d", False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    If http.Status = 200 Then resp = http.responseText
    On Error GoTo 0
    If resp = "" Then Exit Function
    Dim key As String: key = Chr(34) & "regularMarketPrice" & Chr(34) & ":"
    Dim p As Long: p = InStr(resp, key)
    If p = 0 Then Exit Function
    p = p + Len(key)
    Dim q As Long: q = p
    Do While q <= Len(resp)
        If InStr("0123456789.", Mid(resp, q, 1)) = 0 Then Exit Do
        q = q + 1
    Loop
    If q > p Then FetchLiveFx = Val(Mid(resp, p, q - p))
End Function

' USD/TWD from the page's input cell (fallback when the live fetch fails).
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
    ws.cells(r, 1).Font.Color = RGB(0, 200, 255)
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
        Call NavNotify("HC!: need at least 2 holdings", True)
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
        Call NavNotify("HC!: only " & usedCount & " common trading days - per-ticker counts in the Immediate window", True)
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
    Call NavNotify("HC! done " & Format(Now, "hh:mm:ss") & " - holdings correlation updated")
End Sub

' ------------------------------------------------------------
Private Sub HoldingsCorrRenderSheet(tickers() As String, tickerCount As Long, _
                                     corrMatrix() As Double, usedCount As Long, _
                                     commonDates() As Date)
    ' v2 (2026-09-12): RR4 page look - orange accent, grey dividers, a
    ' continuous red/green heat map instead of 7 bins, a colour ramp legend,
    ' AVG CORR per holding, the most / least correlated pairs, and a
    ' portfolio-wide average pairwise correlation.
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
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
    wsC.Activate
    ActiveWindow.DisplayGridlines = False
    Dim rr As Long
    For rr = 1 To 60: wsC.Rows(rr).RowHeight = 18: Next rr

    Dim retCount As Long: retCount = usedCount - 1
    Dim i As Long, j As Long

    ' ---- header --------------------------------------------------
    With wsC.cells(1, 1)
        .Value = "HOLDINGS CORRELATION"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 14
    End With
    wsC.Rows(1).RowHeight = 24
    With wsC.cells(1, 4)
        .Value = retCount & "-day log returns  .  " & Format(commonDates(0), "yyyy/m/d") & " ~ " & _
                 Format(commonDates(usedCount - 1), "yyyy/m/d") & "  .  updated " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
    End With
    Dim lastCol As Long: lastCol = tickerCount + 3       ' matrix cols 2..n+1, gap, AVG at n+3
    With wsC.Range(wsC.cells(1, 1), wsC.cells(1, lastCol + 6)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    ' ---- per-holding average correlation + portfolio average ----------
    Dim avgCorr() As Double: ReDim avgCorr(0 To tickerCount - 1)
    For i = 0 To tickerCount - 1
        Dim sRow As Double: sRow = 0
        For j = 0 To tickerCount - 1
            If i <> j Then sRow = sRow + corrMatrix(i, j)
        Next j
        If tickerCount > 1 Then avgCorr(i) = sRow / (tickerCount - 1)
    Next i

    ' 2026-09-18: the portfolio figure is WEIGHT-averaged, not a plain mean.
    ' A plain mean lets a 0.2%-of-book holding move the headline number as
    ' much as a 20% one, so it read "well spread" while the weight actually
    ' sat in a few correlated names. Each pair is weighted by w_i*w_j, the
    ' same product that drives portfolio variance:
    '     sum(w_i*w_j*corr_ij) / sum(w_i*w_j)   over i<j
    ' Weights come from the RR4 page's WT% column (summed across brokers),
    ' so the number matches the weights on screen. With no usable weights
    ' (RR4 page not built yet) it falls back to equal weights, which makes
    ' this identical to the old plain mean.
    Dim wmap As Object: Set wmap = RR4WeightMap()
    Dim wArr() As Double: ReDim wArr(0 To tickerCount - 1)
    Dim wSum As Double: wSum = 0
    For i = 0 To tickerCount - 1
        Dim wk As String: wk = UCase(Trim(tickers(i)))
        If wmap.Exists(wk) Then wArr(i) = wmap(wk) Else wArr(i) = 0
        If wArr(i) < 0 Then wArr(i) = 0              ' no shorts on this page
        wSum = wSum + wArr(i)
    Next i
    If wSum <= 0 Then
        For i = 0 To tickerCount - 1: wArr(i) = 1: Next i
    End If

    Dim wNum As Double, wDen As Double
    For i = 0 To tickerCount - 1
        For j = i + 1 To tickerCount - 1
            Dim pw As Double: pw = wArr(i) * wArr(j)
            wNum = wNum + pw * corrMatrix(i, j)
            wDen = wDen + pw
        Next j
    Next i
    Dim portAvg As Double: If wDen > 0 Then portAvg = wNum / wDen

    ' ---- matrix --------------------------------------------------
    Dim HDR As Long: HDR = 3
    wsC.Columns(1).ColumnWidth = 12
    For i = 0 To tickerCount - 1
        wsC.Columns(i + 2).ColumnWidth = 8
    Next i
    wsC.Columns(tickerCount + 2).ColumnWidth = 2
    wsC.Columns(lastCol).ColumnWidth = 10

    With wsC.cells(HDR, 1)
        .Value = "TICKER"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter
    End With
    For i = 0 To tickerCount - 1
        With wsC.cells(HDR, i + 2)
            .Value = ShortTicker(tickers(i))
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .Font.Size = 8
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next i
    With wsC.cells(HDR, lastCol)
        .Value = "AVG CORR"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 8
        .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter
    End With
    wsC.Rows(HDR).RowHeight = 20
    With wsC.Range(wsC.cells(HDR, 1), wsC.cells(HDR, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    Dim minAvg As Double: minAvg = 2
    Dim maxAvg As Double: maxAvg = -2
    For i = 0 To tickerCount - 1
        If avgCorr(i) < minAvg Then minAvg = avgCorr(i)
        If avgCorr(i) > maxAvg Then maxAvg = avgCorr(i)
    Next i

    For i = 0 To tickerCount - 1
        Dim rn As Long: rn = HDR + 1 + i
        wsC.Rows(rn).RowHeight = 22
        With wsC.cells(rn, 1)
            .Value = ShortTicker(tickers(i))
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(10, 10, 10)
        End With
        For j = 0 To tickerCount - 1
            With wsC.cells(rn, j + 2)
                .HorizontalAlignment = xlCenter
                .Font.Size = 8
                If i = j Then
                    .Value = ChrW(&H2014)
                    .Interior.Color = RGB(20, 20, 20)
                    .Font.Color = RGB(70, 70, 70)
                Else
                    .Value = corrMatrix(i, j)
                    .NumberFormat = "0.00"
                    .Font.Bold = (Abs(corrMatrix(i, j)) >= 0.5)
                    .Interior.Color = CorrHeatBg(corrMatrix(i, j))
                    .Font.Color = CorrHeatFg(corrMatrix(i, j))
                End If
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = RGB(0, 0, 0)
                    .Weight = xlThin
                End With
            End With
        Next j
        ' AVG CORR column: the most diversifying holding (lowest) in green,
        ' the most crowded (highest) in red
        With wsC.cells(rn, lastCol)
            .Value = avgCorr(i)
            .NumberFormat = "0.00"
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(12, 12, 12)
            If tickerCount > 2 And avgCorr(i) = minAvg Then
                .Font.Color = RGB(0, 200, 90)
            ElseIf tickerCount > 2 And avgCorr(i) = maxAvg Then
                .Font.Color = RGB(235, 70, 70)
            Else
                .Font.Color = RGB(200, 200, 200)
            End If
        End With
    Next i
    Dim lastMatrixRow As Long: lastMatrixRow = HDR + tickerCount
    With wsC.Range(wsC.cells(lastMatrixRow, 1), wsC.cells(lastMatrixRow, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    ' ---- legend: colour ramp + portfolio average ----------------------
    Dim r As Long: r = lastMatrixRow + 2
    wsC.cells(r, 1).Value = "SCALE"
    wsC.cells(r, 1).Font.Color = RR4_ACCENT
    wsC.cells(r, 1).Font.Bold = True
    Dim ramp As Variant: ramp = Array(-1, -0.75, -0.5, -0.25, 0, 0.25, 0.5, 0.75, 1)
    Dim k As Long
    For k = 0 To UBound(ramp)
        ' the ramp may run into the narrow gap column - widen it enough to show
        If wsC.Columns(k + 2).ColumnWidth < 7 Then wsC.Columns(k + 2).ColumnWidth = 7
        With wsC.cells(r, k + 2)
            .Value = ramp(k)
            .NumberFormat = "+0.00;-0.00;0.00"
            .Font.Size = 8
            .HorizontalAlignment = xlCenter
            .Interior.Color = CorrHeatBg(CDbl(ramp(k)))
            .Font.Color = CorrHeatFg(CDbl(ramp(k)))
        End With
    Next k
    With wsC.cells(r, UBound(ramp) + 3)
        .Value = "red = move together   green = move against   (Pearson on daily log returns, common window)"
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
    End With
    r = r + 1
    wsC.cells(r, 1).Value = "PORTFOLIO WEIGHT-AVG PAIRWISE CORR"
    wsC.cells(r, 1).Font.Color = RGB(0, 200, 255)
    With wsC.cells(r, 4)
        .Value = portAvg
        .NumberFormat = "0.00"
        .Font.Bold = True
        .Font.Color = IIf(portAvg >= 0.5, RGB(235, 70, 70), IIf(portAvg <= 0.2, RGB(0, 200, 90), RGB(221, 221, 221)))
    End With
    wsC.cells(r, 5).Value = IIf(portAvg >= 0.5, "crowded - holdings largely share one bet", _
                            IIf(portAvg <= 0.2, "well spread", "moderate"))
    wsC.cells(r, 5).Font.Color = RGB(120, 120, 120)
    wsC.cells(r, 5).Font.Size = 9

    ' ---- top / bottom pairs -------------------------------------
    r = r + 2
    Dim nP As Long: nP = tickerCount * (tickerCount - 1) \ 2
    Dim pA() As Long, pB() As Long, pV() As Double
    ReDim pA(1 To nP): ReDim pB(1 To nP): ReDim pV(1 To nP)
    k = 0
    For i = 0 To tickerCount - 1
        For j = i + 1 To tickerCount - 1
            k = k + 1: pA(k) = i: pB(k) = j: pV(k) = corrMatrix(i, j)
        Next j
    Next i
    ' simple selection sort, descending
    Dim a As Long, b As Long, tmpL As Long, tmpD As Double
    For a = 1 To nP - 1
        For b = a + 1 To nP
            If pV(b) > pV(a) Then
                tmpD = pV(a): pV(a) = pV(b): pV(b) = tmpD
                tmpL = pA(a): pA(a) = pA(b): pA(b) = tmpL
                tmpL = pB(a): pB(a) = pB(b): pB(b) = tmpL
            End If
        Next b
    Next a
    Dim show As Long: show = IIf(nP < 5, nP, 5)

    wsC.cells(r, 1).Value = "MOST CORRELATED PAIRS"
    wsC.cells(r, 1).Font.Color = RR4_ACCENT
    wsC.cells(r, 1).Font.Bold = True
    wsC.cells(r, 6).Value = "LEAST CORRELATED / HEDGING PAIRS"
    wsC.cells(r, 6).Font.Color = RR4_ACCENT
    wsC.cells(r, 6).Font.Bold = True
    r = r + 1
    For k = 1 To show
        Call CorrPairLine(wsC, r + k - 1, 1, ShortTicker(tickers(pA(k))), ShortTicker(tickers(pB(k))), pV(k))
        Dim kk As Long: kk = nP - k + 1
        Call CorrPairLine(wsC, r + k - 1, 6, ShortTicker(tickers(pA(kk))), ShortTicker(tickers(pB(kk))), pV(kk))
    Next k
    r = r + show + 1

    ' ---- footnote -----------------------------------------------
    Dim notes As Variant
    notes = Array( _
        "HOW THESE ARE COMPUTED", _
        "CORRELATION   = Pearson correlation of daily LOG returns over the dates every holding has in common (" & retCount & " returns)", _
        "AVG CORR      = mean of a holding's correlations with the other holdings   - lowest (green) diversifies most, highest (red) is the most crowded", _
        "PORTFOLIO WEIGHT-AVG = sum(w_i*w_j*corr_ij) / sum(w_i*w_j) over each pair once, w = WT% on the RR4 page   - >= 0.5 crowded, <= 0.2 well spread", _
        "PAIRS         = the same matrix ranked: top 5 pairs that move together, bottom 5 that move against each other")
    For k = 0 To UBound(notes)
        With wsC.cells(r + k, 1)
            .Value = notes(k)
            .Font.Size = 9
            If k = 0 Then
                .Font.Color = RR4_ACCENT: .Font.Bold = True
            Else
                .Font.Color = RGB(120, 120, 120)
            End If
        End With
    Next k

    Call NavAdd(wsC, "HC")
    wsC.Activate
End Sub

' Ticker (as logged, upper-cased) -> its WT% on the RR4 page's position log,
' summed over brokers because that page keeps one row per Ticker|Broker while
' the correlation matrix has one row per ticker. Empty when the page has no
' positions on it yet - callers fall back to equal weights.
Private Function RR4WeightMap() As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(SH_PORT): On Error GoTo 0
    If Not ws Is Nothing Then
        Dim lastR As Long: lastR = LastPositionRow(ws)
        Dim r As Long
        For r = RR4_POS_FIRST To lastR
            Dim tk As String: tk = UCase(Trim(CellStr(ws.cells(r, RR4_LEFT + 1).Value)))
            If tk <> "" Then m(tk) = m(tk) + NumOr0(ws.cells(r, RR4_LEFT + 12).Value)
        Next r
    End If
    Set RR4WeightMap = m
End Function

' "3653 x 7610    0.82" with the value in the heat colour
Private Sub CorrPairLine(ws As Worksheet, r As Long, c As Long, t1 As String, t2 As String, v As Double)
    With ws.cells(r, c)
        .Value = t1 & " x " & t2
        .Font.Color = RGB(200, 200, 200)
    End With
    With ws.cells(r, c + 3)
        .Value = v
        .NumberFormat = "0.00"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = CorrHeatBg(v)
        .Font.Color = CorrHeatFg(v)
    End With
End Sub

' Continuous heat: black at 0, towards red for +1 and green for -1.
Public Function CorrHeatBg(ByVal v As Double) As Long
    Dim t As Double: t = Abs(v): If t > 1 Then t = 1
    If v >= 0 Then
        CorrHeatBg = RGB(CLng(14 + (170 - 14) * t), CLng(14 + (40 - 14) * t), CLng(14 + (25 - 14) * t))
    Else
        CorrHeatBg = RGB(CLng(14 + (0 - 14) * t), CLng(14 + (120 - 14) * t), CLng(14 + (55 - 14) * t))
    End If
End Function

Public Function CorrHeatFg(ByVal v As Double) As Long
    If Abs(v) >= 0.5 Then
        CorrHeatFg = RGB(255, 255, 255)
    ElseIf Abs(v) >= 0.2 Then
        CorrHeatFg = RGB(220, 220, 220)
    Else
        CorrHeatFg = RGB(140, 140, 140)
    End If
End Function

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
    ' risk contribution (v4.12): w_i * cov(r_i, r_p) / var(r_p) - the share of
    ' the portfolio's variance each holding is responsible for; sums to 100%
    ' (weights sum to 1 and every series runs over the same common window)
    Dim riskContrib() As Double: ReDim riskContrib(0 To n - 1)
    If retCount >= 2 Then
        Dim portRet() As Double
        ReDim portRet(0 To retCount - 1)
        Dim stkRet() As Double
        ReDim stkRet(0 To n - 1, 0 To retCount - 1)

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
                    stkRet(i, j) = (p1 - p0) / p0
                    wsum = wsum + (mktVals(i) / totalMktTWD) * stkRet(i, j)
                End If
            Next i
            portRet(j) = wsum
        Next j

        portDaysUsed = retCount
        On Error Resume Next
        portAnnVol = Application.WorksheetFunction.StDev_S(portRet) * Sqr(252)
        On Error GoTo 0

        ' covariance of each holding with the portfolio
        Dim meanP As Double
        For j = 0 To retCount - 1: meanP = meanP + portRet(j): Next j
        meanP = meanP / retCount
        Dim varP As Double
        For j = 0 To retCount - 1: varP = varP + (portRet(j) - meanP) ^ 2: Next j
        varP = varP / (retCount - 1)
        If varP > 0 And totalMktTWD > 0 Then
            For i = 0 To n - 1
                Dim meanI As Double: meanI = 0
                For j = 0 To retCount - 1: meanI = meanI + stkRet(i, j): Next j
                meanI = meanI / retCount
                Dim covIP As Double: covIP = 0
                For j = 0 To retCount - 1: covIP = covIP + (stkRet(i, j) - meanI) * (portRet(j) - meanP): Next j
                covIP = covIP / (retCount - 1)
                riskContrib(i) = (mktVals(i) / totalMktTWD) * covIP / varP
            Next i
        End If
    End If

    ' ---- 4. Render ----
    Call VolRenderSheet(tickers, mktVals, totalMktTWD, n, stockAnnVol, stockDaysUsed, _
                         portAnnVol, portDaysUsed, commonDates, usedCount, riskContrib)

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
                            commonDates() As Date, usedCount As Long, riskContrib() As Double)
    Dim wsV As Worksheet
    On Error Resume Next
    Set wsV = ThisWorkbook.Sheets("Volatility180D")
    On Error GoTo 0
    If wsV Is Nothing Then
        Set wsV = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsV.Name = "Volatility180D"
    End If

    ' v4.12 (2026-09-12): RR4 page look - orange accent, grey dividers,
    ' 8/14 stripes, muted labels; plus RISK CONTRIB%, diversification ratio
    ' and a 1-day 95% VaR line.
    Call NavStrip(wsV)
    wsV.Cells.Clear
    With wsV.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    wsV.Activate
    ActiveWindow.DisplayGridlines = False
    Dim rr As Long
    For rr = 1 To 60: wsV.Rows(rr).RowHeight = 18: Next rr

    With wsV.Cells(1, 1)
        .Value = "PORTFOLIO 180D VOLATILITY"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 14
    End With
    wsV.Rows(1).RowHeight = 24
    With wsV.Cells(1, 4)
        .Value = "updated " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
    End With
    With wsV.Range(wsV.Cells(1, 1), wsV.Cells(1, 7)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    Dim r As Long: r = 3
    wsV.Cells(r, 1).Value = "PORTFOLIO (MARKET-VALUE WEIGHTED)"
    wsV.Cells(r, 1).Font.Color = RR4_ACCENT
    wsV.Cells(r, 1).Font.Bold = True
    r = r + 1

    Dim portDailyVol As Double: If portAnnVol > 0 Then portDailyVol = portAnnVol / Sqr(252)
    Dim sumWVol As Double, i As Long
    For i = 0 To n - 1
        If totalMktTWD > 0 Then sumWVol = sumWVol + mktVals(i) / totalMktTWD * stockAnnVol(i)
    Next i
    Dim divRatio As Double: If portAnnVol > 0 Then divRatio = sumWVol / portAnnVol
    Dim var95 As Double: var95 = 1.645 * portDailyVol * totalMktTWD

    ' left column
    Call WriteKV(wsV, r, "ANNUALIZED STD DEV", portAnnVol, "0.00%", False)
    Call WriteKV(wsV, r + 1, "DAILY STD DEV", portDailyVol, "0.00%", False)
    Call WriteKV(wsV, r + 2, "1-DAY VAR 95% (TWD)", -var95, "#,##0", True)
    wsV.Cells(r + 3, 1).Value = "TRADING DAYS USED"
    wsV.Cells(r + 3, 1).Font.Color = RGB(0, 200, 255)
    wsV.Cells(r + 3, 2).Value = portDaysUsed & " / 180"
    wsV.Cells(r + 3, 2).Font.Color = IIf(portDaysUsed < 180, RR4_ACCENT, RGB(221, 221, 221))
    wsV.Cells(r + 3, 2).Font.Bold = True
    ' right column
    Call WriteKV2(wsV, r, 4, "WEIGHTED AVG STDEV", sumWVol, "0.00%")
    Call WriteKV2(wsV, r + 1, 4, "DIVERSIFICATION RATIO", divRatio, "0.00x")
    wsV.Cells(r + 1, 6).Value = IIf(divRatio >= 1.5, "well diversified", IIf(divRatio >= 1.2, "moderate", "concentrated"))
    wsV.Cells(r + 1, 6).Font.Color = RGB(120, 120, 120)
    wsV.Cells(r + 1, 6).Font.Size = 9
    If usedCount >= 2 Then
        wsV.Cells(r + 2, 4).Value = "COMMON DATE RANGE"
        wsV.Cells(r + 2, 4).Font.Color = RGB(0, 200, 255)
        wsV.Cells(r + 2, 5).Value = Format(commonDates(0), "yyyy/m/d") & " ~ " & _
                                     Format(commonDates(usedCount - 1), "yyyy/m/d")
        wsV.Cells(r + 2, 5).Font.Color = RGB(221, 221, 221)
    End If
    r = r + 4
    If portDaysUsed > 0 And portDaysUsed < 180 Then
        wsV.Cells(r, 1).Value = "* shortest-history holding capped the common window below 180 days"
        wsV.Cells(r, 1).Font.Color = RGB(120, 120, 120)
        wsV.Cells(r, 1).Font.Italic = True
        wsV.Cells(r, 1).Font.Size = 9
        r = r + 1
    End If

    r = r + 1
    wsV.Cells(r, 1).Value = "PER-STOCK 180D STANDARD DEVIATION"
    wsV.Cells(r, 1).Font.Color = RR4_ACCENT
    wsV.Cells(r, 1).Font.Bold = True
    With wsV.Cells(r, 4)
        .Value = "RISK CONTRIB% = weight x cov(stock, portfolio) / var(portfolio)  -  sums to 100%"
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
    End With
    r = r + 1

    Dim hdrs As Variant
    hdrs = Array("TICKER", "WEIGHT%", "DAYS USED", "DAILY STDEV", "ANNUALIZED STDEV", "RISK CONTRIB%", "RISK/WEIGHT")
    Dim c As Long
    For c = 0 To UBound(hdrs)
        With wsV.Cells(r, c + 1)
            .Value = hdrs(c)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .Font.Size = 9
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next c
    With wsV.Range(wsV.Cells(r, 1), wsV.Cells(r, 7)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With
    wsV.Rows(r).RowHeight = 20
    r = r + 1

    Dim maxContrib As Double
    For i = 0 To n - 1
        If riskContrib(i) > maxContrib Then maxContrib = riskContrib(i)
    Next i

    Dim firstDataRow As Long: firstDataRow = r
    For i = 0 To n - 1
        Dim rowBg As Long: rowBg = IIf(i Mod 2 = 0, RGB(8, 8, 8), RGB(14, 14, 14))
        With wsV.Range(wsV.Cells(r, 1), wsV.Cells(r, 7))
            .Interior.Color = rowBg
            .HorizontalAlignment = xlCenter
        End With
        wsV.Rows(r).RowHeight = 22

        wsV.Cells(r, 1).Value = tickers(i)
        wsV.Cells(r, 1).Font.Color = RR4_ACCENT
        wsV.Cells(r, 1).Font.Bold = True

        Dim w As Double: w = 0
        If totalMktTWD > 0 Then w = mktVals(i) / totalMktTWD
        wsV.Cells(r, 2).Value = w
        wsV.Cells(r, 2).NumberFormat = "0.00%"

        wsV.Cells(r, 3).Value = stockDaysUsed(i) & IIf(stockDaysUsed(i) < 180, "*", "")
        wsV.Cells(r, 3).Font.Color = IIf(stockDaysUsed(i) < 180, RR4_ACCENT, RGB(160, 160, 160))

        Dim dailyVol As Double: dailyVol = 0
        If stockAnnVol(i) > 0 Then dailyVol = stockAnnVol(i) / Sqr(252)
        wsV.Cells(r, 4).Value = dailyVol
        wsV.Cells(r, 4).NumberFormat = "0.00%"

        wsV.Cells(r, 5).Value = stockAnnVol(i)
        wsV.Cells(r, 5).NumberFormat = "0.00%"
        wsV.Cells(r, 5).Font.Bold = True

        wsV.Cells(r, 6).Value = riskContrib(i)
        wsV.Cells(r, 6).NumberFormat = "0.0%"
        wsV.Cells(r, 6).Font.Bold = True
        ' the single biggest risk contributor is lit
        If maxContrib > 0 And riskContrib(i) = maxContrib Then
            wsV.Cells(r, 6).Font.Color = RR4_ACCENT
        End If

        ' risk share vs weight share: >1 = punches above its weight
        If w > 0 Then
            wsV.Cells(r, 7).Value = riskContrib(i) / w
            wsV.Cells(r, 7).NumberFormat = "0.00x"
            wsV.Cells(r, 7).Font.Color = IIf(riskContrib(i) / w > 1.3, RGB(235, 70, 70), _
                                        IIf(riskContrib(i) / w < 0.7, RGB(0, 200, 90), RGB(200, 200, 200)))
        End If
        r = r + 1
    Next i
    With wsV.Range(wsV.Cells(r - 1, 1), wsV.Cells(r - 1, 7)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    r = r + 1
    wsV.Cells(r, 1).Value = "* fewer than 180 trading days of history available - used actual days shown"
    wsV.Cells(r, 1).Font.Color = RGB(120, 120, 120)
    wsV.Cells(r, 1).Font.Italic = True
    wsV.Cells(r, 1).Font.Size = 9
    r = r + 2
    ' how each number is computed (v4.12.1)
    wsV.Cells(r, 1).Value = "HOW THESE ARE COMPUTED"
    wsV.Cells(r, 1).Font.Color = RR4_ACCENT
    wsV.Cells(r, 1).Font.Bold = True
    wsV.Cells(r, 1).Font.Size = 9
    r = r + 1
    Dim notes As Variant
    notes = Array( _
        "DAILY STDEV          = sample std dev of the last 180 daily returns (each stock on its own most-recent window)", _
        "ANNUALIZED STDEV     = DAILY STDEV x sqrt(252)", _
        "PORTFOLIO STD DEV    = std dev of the market-value-weighted daily return, all stocks on the COMMON date window", _
        "WEIGHTED AVG STDEV   = sum( weight_i x annualized stdev_i )   - what the portfolio vol would be if every holding moved together", _
        "DIVERSIFICATION RATIO= WEIGHTED AVG STDEV / PORTFOLIO STD DEV   - 1.0x = no diversification benefit; higher = lower correlation between holdings", _
        "1-DAY VAR 95%        = 1.645 x portfolio DAILY STD DEV x total market value   - normal-distribution assumption; 1 day in 20 should lose more", _
        "RISK CONTRIB%        = weight_i x cov(return_i, portfolio return) / var(portfolio return)   - share of portfolio variance; the column sums to 100%", _
        "RISK/WEIGHT          = RISK CONTRIB% / WEIGHT%   - > 1.3 (red) adds more risk than its size, < 0.7 (green) dampens the book")
    Dim ni As Long
    For ni = 0 To UBound(notes)
        With wsV.Cells(r, 1)
            .Value = notes(ni)
            .Font.Color = RGB(120, 120, 120)
            .Font.Size = 9
        End With
        r = r + 1
    Next ni

    wsV.Columns(1).ColumnWidth = 24
    wsV.Columns(2).ColumnWidth = 12
    wsV.Columns(3).ColumnWidth = 12
    wsV.Columns(4).ColumnWidth = 22
    wsV.Columns(5).ColumnWidth = 18
    wsV.Columns(6).ColumnWidth = 15
    wsV.Columns(7).ColumnWidth = 13
    Call NavAdd(wsV, "V")
    ThisWorkbook.Sheets(SH_PORT).Activate
End Sub

' label in column c, value in c+1 (used for the right-hand KV column)
Private Sub WriteKV2(ws As Worksheet, r As Long, c As Long, label As String, _
                     val As Double, fmt As String)
    ws.cells(r, c).Value = label
    ws.cells(r, c).Font.Color = RGB(0, 200, 255)
    ws.cells(r, c + 1).Value = val
    ws.cells(r, c + 1).NumberFormat = fmt
    ws.cells(r, c + 1).Font.Bold = True
    ws.cells(r, c + 1).Font.Color = RGB(221, 221, 221)
End Sub

' v4.13: a sheet still on the v4.12 layout has the bar's "P" badge in B1
' (with the blank row it sits in B2).  Inserting one row at the top moves
' EVERYTHING - hand-typed cells, the config cells T1/T2 -> T2/T3, charts
' (xlMove), the watchlist, the hidden order column - onto the new row
' numbers in one step, so the read-before-clear code finds it all.
Public Sub MigrateRR4TopRow(ws As Worksheet)
    If UCase(Trim(CStr(ws.cells(1, RR4_LEFT + 1).Value))) <> "P" Then Exit Sub
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    Call ShapesMoveOnly(ws)
    ws.Rows(1).Insert Shift:=xlDown
    ws.Rows(1).ClearFormats
    ws.Rows(1).Interior.Color = RGB(0, 0, 0)
    ws.Rows(1).RowHeight = 14
    ' names that pointed at T1/T2 followed the insert; repoint by constant anyway
    ThisWorkbook.names.Add "InceptionDate", ws.Range(RR4_CFG_INC)
    ThisWorkbook.names.Add "StartingCapital", ws.Range(RR4_CFG_CAP)
    On Error GoTo 0
    Application.EnableEvents = prevEv
End Sub
