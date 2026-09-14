Attribute VB_Name = "TickerInsight"
Option Explicit

' ================================================================
' TICKER INSIGHT v2.0 - ticker panel on the RR4 page
' Per-ticker FIFO trade history & performance, Bloomberg style,
' US colour convention (green +, red -)
'
' v2 (2026-09-11): the separate "Ticker Insight" sheet is gone. The
' panel now lives in the upper-right block of the RR4 sheet (J4:S24),
' next to the summary; PortfolioDashboard_v3 owns the rest of the page
' and the column widths / row heights. Nothing here may delete shapes or
' touch cells outside J4:S24 (plus the hidden tracker X1) - the weight
' bar, the weight donut and the nav bar (rows 1-3) share the sheet.
' v2.1 (2026-09-12): everything moved down RR4_TOP = 3 rows for the nav
' bar; the row numbers in the layout below are the v2 ones + 3.
' v2.2 (2026-09-12): and one column right (RR4_LEFT = 1, column A is now a
' blank spacer), so the panel sits in J:S and its cells are K4 / L20.
' v2.4 (2026-09-13): ResolveTicker - a bare TW code ("3374") takes the
' .TW/.TWO suffix of its Transactions rows before any FX decision; typed
' without the suffix it used to be classified US and every TWD aggregate
' was multiplied by USD/TWD.
'
' ENTRY POINTS:
'   Sub OpenTickerInsight()            - jump to the panel's ticker cell
'   Sub RefreshTickerInsight(ticker)   - full rebuild for a ticker
'                                        (blank ticker = empty shell)
'   Sub RefreshTickerProjection        - recompute PROJECTION block only
'   Sub RenderTickerPanel(t, target)   - used by RebuildPortfolioDashboard
'                                        to redraw after it clears the page
'   Sub TickerInsight_OnChange(t)      - dispatcher for sheet events
'
' PANEL LAYOUT (RR4 sheet):
'   (sheet rows; the page rows they came from are these minus RR4_TOP = 3)
'   J5 TICKER <GO> label, J6 [ticker input], K6 company name
'   (v2.5, 2026-09-14: stacked label-over-input like USD/TWD / ARRANGE; was J5 / K5 / L5)
'   J8  1) ACTIVE POSITION      J9:J13  labels, L9:L13  values
'   J15 2) LIFETIME METRICS     J16:J18 labels, L16:L18 values
'   J20 3) PROJECTION           J21:J24 labels, L21:L24 values
'                               L22 = PRICE TARGET [input]
'   N8  4) RECENT TRADE HISTORY N9:S9 header, N10:S24 newest first,
'                               N25 "+n older" when it does not fit
'   (v2.6, 2026-09-14: the body sits one row lower than v2.5 - rows 8-24)
'   X1  (hidden) ticker of the last render - keeps L20 when the same
'       ticker is refreshed, clears it when the ticker changes
'
' The RR4 sheet's Worksheet_Change calls RefreshTickerInsight on K4 and
' RefreshTickerProjection on L20 (see RR4/SheetRR4_Code.txt).
' ================================================================

Private Const SH_RR4 As String = "RR4"
Private Const SH_TR  As String = "Transactions"

Public Const TI_TICKER_CELL As String = "J6"   ' v2.5: under the TICKER <GO> label (was K5)
Public Const TI_TARGET_CELL As String = "L22"   ' v2.6: body shifted one row down (was L21)
Private Const TI_TRACK_CELL As String = "X2"   ' clear of the page and the config cells (row 1 stays blank, v4.13)

Private Const TI_LBL    As Long = 10    ' J - labels
Private Const TI_VAL    As Long = 12    ' L - values
Private Const TI_HCOL   As Long = 14    ' N - first trade-history column
Private Const TI_BODY   As Long = 1     ' v2.6 (2026-09-14): sections 1-4 start one row lower (row 8, ending row 24)
Private Const TI_HHDR   As Long = 9     ' trade-history header row (v4.13: 8; v2.6: 9)
Private Const TI_H1     As Long = 10    ' first trade-history row
Private Const TI_HLAST  As Long = 24    ' last trade-history row (v2.3: was 25; v4.13: 24)
Private Const TI_BOTTOM As Long = 25    ' last panel row (the chart band starts at RR4_CHART_TOP = 27)
Private Const TI_RIGHT  As Long = 19    ' S - last panel column

' Transactions PAGE columns (A=1 ..); the sheet column is TrCol(wsTr, n)
' and the first data row TrHdrRow(wsTr) + 1 - the page carries the nav
' bar since 2026-09-13 (Attach.bas geometry helpers).
Private Const COL_DATE   As Long = 2
Private Const COL_TICKER As Long = 3
Private Const COL_ACTION As Long = 4
Private Const COL_SHARES As Long = 5
Private Const COL_NETAMT As Long = 9
Private Const COL_BROKER As Long = 14

' Fallback USD -> TWD rate when RR4!B2 holds no usable rate. All aggregate
' PnL / exposure values for US tickers are converted; per-share prices
' (Entry, Last, Price Target) stay in NATIVE currency.
Private Const FX_USD_TWD As Double = 31.6

' ----------------------------------------------------------------
' One FIFO-paired trade record (one BUY lot matched to one SELL)
' ----------------------------------------------------------------
Private Type TIRecord
    EntryDate As Date
    ExitDate  As Date
    DaysHeld  As Long
    shares    As Double
    Cost      As Double
    Proceeds  As Double
    PnL       As Double
    ReturnPct As Double
    broker    As String
End Type

' ----------------------------------------------------------------
' One residual lot still open after all transactions processed
' ----------------------------------------------------------------
Private Type TILot
    lotDate      As Date
    shares       As Double
    costPerShare As Double
    broker       As String
End Type

' ================================================================
' PUBLIC ENTRY: jump to the panel on RR4
' ================================================================
Public Sub OpenTickerInsight()
    Dim ws As Worksheet: Set ws = PanelSheet()
    ws.Activate
    ws.Range(TI_TICKER_CELL).Select
    Call RefreshTickerInsight(CStr(ws.Range(TI_TICKER_CELL).Value))
End Sub

' ================================================================
' DIAGNOSTIC: dumps the state of the Ticker Insight pipeline for the
' ticker currently in K4 (or a passed-in one). Useful when no values
' appear in the panel. Run via Alt+F8 -> DiagnoseTicker.
' ================================================================
Public Sub DiagnoseTicker(Optional ByVal forcedTicker As String = "")
    Dim msg As String
    Dim lf As String: lf = vbCrLf

    On Error GoTo ErrTrap

    ' --- 1. Sheets ---
    msg = msg & "=== SHEET CHECK ===" & lf
    Dim wsTI As Worksheet, wsTr As Worksheet
    On Error Resume Next
    Set wsTI = ThisWorkbook.Sheets(SH_RR4)
    Set wsTr = ThisWorkbook.Sheets(SH_TR)
    On Error GoTo ErrTrap

    msg = msg & "RR4 (panel)   : " & IIf(wsTI Is Nothing, "MISSING", "OK") & lf
    msg = msg & "Transactions  : " & IIf(wsTr Is Nothing, "MISSING", "OK") & lf

    If wsTr Is Nothing Then GoTo ShowMsg

    ' --- 2. Ticker resolution ---
    msg = msg & lf & "=== TICKER ===" & lf
    Dim ticker As String
    If forcedTicker <> "" Then
        ticker = forcedTicker
    ElseIf Not (wsTI Is Nothing) Then
        ticker = CStr(wsTI.Range(TI_TICKER_CELL).Value)
    End If
    ticker = UCase(Trim(ticker))

    If ticker = "" Then
        msg = msg & "(" & TI_TICKER_CELL & " empty - type a ticker then re-run, or DiagnoseTicker " & _
              Chr(34) & "CRDO" & Chr(34) & ")" & lf
        GoTo ShowMsg
    End If

    msg = msg & "Input         : [" & ticker & "]" & lf
    msg = msg & "Normalized    : [" & NormalizeTicker(ticker) & "]" & lf
    msg = msg & "Is TW ticker  : " & IsTWTicker(ticker) & lf
    msg = msg & "FX multiplier : " & GetFXToTWD(ticker) & lf
    msg = msg & "Live price    : " & GetStockPriceSafe(ticker) & lf

    ' --- 3. Transactions scan ---
    msg = msg & lf & "=== TRANSACTIONS SCAN ===" & lf
    Dim lastRow As Long
    lastRow = TrLastRow(wsTr)
    msg = msg & "Total rows    : " & (lastRow - TrHdrRow(wsTr)) & lf

    Dim normT As String: normT = NormalizeTicker(ticker)
    Dim matchCount As Long, sample As String, samples As Long
    Dim r As Long
    For r = TrHdrRow(wsTr) + 1 To lastRow
        Dim cellT As String: cellT = CStr(wsTr.cells(r, TrCol(wsTr, COL_TICKER)).Value)
        If NormalizeTicker(cellT) = normT Then
            matchCount = matchCount + 1
            If samples < 5 Then
                sample = sample & "  R" & r & ":  " & _
                    CStr(wsTr.cells(r, TrCol(wsTr, COL_DATE)).Value) & "  " & _
                    "[" & cellT & "]  " & _
                    CStr(wsTr.cells(r, TrCol(wsTr, COL_ACTION)).Value) & "  " & _
                    "sh=" & CStr(wsTr.cells(r, TrCol(wsTr, COL_SHARES)).Value) & "  " & _
                    "net=" & CStr(wsTr.cells(r, TrCol(wsTr, COL_NETAMT)).Value) & "  " & _
                    "brk=" & CStr(wsTr.cells(r, TrCol(wsTr, COL_BROKER)).Value) & lf
                samples = samples + 1
            End If
        End If
    Next r

    msg = msg & "Matching rows : " & matchCount & lf
    If matchCount > 0 Then
        msg = msg & lf & "Sample rows (up to 5):" & lf & sample
    Else
        msg = msg & "(none -- check ticker spelling and column C of Transactions)" & lf
        msg = msg & "Hint: Transactions C column normalised forms of first 10 tickers:" & lf
        Dim cnt As Long: cnt = 0
        For r = TrHdrRow(wsTr) + 1 To lastRow
            If cnt < 10 Then
                Dim t2 As String: t2 = CStr(wsTr.cells(r, TrCol(wsTr, COL_TICKER)).Value)
                If t2 <> "" Then
                    msg = msg & "  [" & t2 & "] -> [" & NormalizeTicker(t2) & "]" & lf
                    cnt = cnt + 1
                End If
            End If
        Next r
    End If

    ' --- 4. FIFO build ---
    msg = msg & lf & "=== FIFO RESULT ===" & lf
    Dim records() As TIRecord, recCount As Long
    Dim lots() As TILot, lotCount As Long
    Call BuildFIFOHistory(ticker, records, recCount, lots, lotCount)
    msg = msg & "recCount      : " & recCount & "  (closed FIFO pairs)" & lf
    msg = msg & "lotCount      : " & lotCount & "  (open lots = active pos)" & lf

    If recCount > 0 Then
        msg = msg & lf & "First record:" & lf
        msg = msg & "  Entry=" & records(1).EntryDate & "  Exit=" & records(1).ExitDate & lf
        msg = msg & "  Days=" & records(1).DaysHeld & "  Shares=" & records(1).shares & lf
        msg = msg & "  Cost=" & records(1).Cost & "  Proceeds=" & records(1).Proceeds & lf
        msg = msg & "  PnL=" & records(1).PnL & "  Ret%=" & records(1).ReturnPct & lf
        msg = msg & "  Broker=" & records(1).broker & lf
    End If

    If lotCount > 0 Then
        msg = msg & lf & "First open lot:" & lf
        msg = msg & "  Date=" & lots(1).lotDate & "  Shares=" & lots(1).shares & lf
        msg = msg & "  CPS=" & lots(1).costPerShare & "  Broker=" & lots(1).broker & lf
    End If

    ' --- 5. Format probe ---
    msg = msg & lf & "=== FORMAT PROBE ===" & lf
    msg = msg & "MoneyFmt(NT$) = " & MoneyFmt("NT$") & lf
    msg = msg & "PnLFmt(NT$)   = " & PnLFmt("NT$") & lf

    ' --- 6. DrawLifetimeMetrics dry-run ---
    If Not (wsTI Is Nothing) Then
        msg = msg & lf & "=== DRAW PROBE ===" & lf
        On Error Resume Next
        Err.Clear
        Call DrawLifetimeMetrics(wsTI, ticker, records, recCount)
        If Err.Number <> 0 Then
            msg = msg & "DrawLifetimeMetrics: ERR " & Err.Number & " - " & Err.Description & lf
        Else
            msg = msg & "DrawLifetimeMetrics: OK (no error)" & lf
        End If
        On Error GoTo ErrTrap
    End If

ShowMsg:
    MsgBox msg, vbInformation, "Ticker Insight Diagnostic"
    Exit Sub

ErrTrap:
    msg = msg & lf & lf & "*** DIAGNOSTIC EXCEPTION ***" & lf & _
          "Err " & Err.Number & ": " & Err.Description
    MsgBox msg, vbCritical, "Ticker Insight Diagnostic - error"
End Sub

' ================================================================
' DISPATCHER: for callers that route sheet events by address
' ================================================================
Public Sub TickerInsight_OnChange(ByVal Target As Range)
    On Error Resume Next
    If Target.CountLarge > 1 Then Exit Sub
    If Target.Worksheet.Name <> SH_RR4 Then Exit Sub
    If Not Intersect(Target, Target.Worksheet.Range(TI_TICKER_CELL)) Is Nothing Then
        Call RefreshTickerInsight(CStr(Target.Value))
    ElseIf Not Intersect(Target, Target.Worksheet.Range(TI_TARGET_CELL)) Is Nothing Then
        Call RefreshTickerProjection
    End If
End Sub

' ================================================================
' Redraw after RebuildPortfolioDashboard has cleared the page: put the
' saved ticker / target back, then refresh as if the same ticker were
' re-entered (so the target survives).
' ================================================================
Public Sub RenderTickerPanel(ByVal ticker As String, Optional ByVal targetVal As Variant)
    Dim ws As Worksheet: Set ws = PanelSheet()
    ticker = UCase(Trim(ticker))
    ws.Range(TI_TRACK_CELL).Value = ticker
    If Not IsMissing(targetVal) Then
        If Not IsEmpty(targetVal) Then ws.Range(TI_TARGET_CELL).Value = targetVal
    End If
    Call RefreshTickerInsight(ticker)
End Sub

' ================================================================
' FULL REFRESH: rebuild every section for a ticker
' ================================================================
Public Sub RefreshTickerInsight(ByVal ticker As String)
    ticker = ResolveTicker(ticker)
    Dim ws As Worksheet: Set ws = PanelSheet()

    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo ErrTrap

    ' --- Detect ticker change via the hidden T1 tracker ----------
    ' X1 stores the PREVIOUS ticker. K4 has already been updated by the
    ' user, so it can't be used for the comparison.
    Dim oldTicker As String
    oldTicker = UCase(Trim(CStr(ws.Range(TI_TRACK_CELL).Value)))
    Dim savedTarget As Variant: savedTarget = ws.Range(TI_TARGET_CELL).Value

    Call ClearTickerData(ws)
    Call DrawShellHeaders(ws, ticker)
    If ticker = "" Then GoTo StampTicker

    ' If user is refreshing the SAME ticker, preserve PRICE TARGET.
    ' If they switched tickers, leave L20 empty (set by shell).
    If oldTicker = ticker And Not IsError(savedTarget) Then
        If IsNumeric(savedTarget) And Not IsEmpty(savedTarget) Then
            If CDbl(savedTarget) > 0 Then ws.Range(TI_TARGET_CELL).Value = CDbl(savedTarget)
        End If
    End If

    Dim records() As TIRecord, recCount As Long
    Dim lots() As TILot, lotCount As Long
    Call BuildFIFOHistory(ticker, records, recCount, lots, lotCount)

    If recCount = 0 And lotCount = 0 Then
        With ws.cells(RR4_TOP + TI_BODY + 4, TI_LBL)
            .Value = "No transactions found for [" & ticker & "]"
            .Font.Color = RGB(255, 80, 80)
            .Font.Bold = True
            .Font.Italic = True
        End With
        GoTo StampTicker
    End If

    Call DrawActivePosition(ws, ticker, lots, lotCount)
    Call DrawLifetimeMetrics(ws, ticker, records, recCount)
    Call DrawProjection(ws, ticker, lots, lotCount)
    Call DrawTradeHistory(ws, ticker, records, recCount)

StampTicker:
    ws.Range(TI_TRACK_CELL).Value = ticker

Done:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScr
    Exit Sub

ErrTrap:
    ' Surface errors instead of silently swallowing - vital for debug
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScr
    If Application.Visible Then
        MsgBox "RefreshTickerInsight error" & vbCrLf & vbCrLf & _
               "Err " & errN & ": " & errD & vbCrLf & _
               "Ticker: " & ticker, vbCritical, "Ticker Insight"
    End If
End Sub

' ================================================================
' LIGHT REFRESH: recompute PROJECTION only (price target changed)
' ================================================================
Public Sub RefreshTickerProjection()
    Dim ws As Worksheet: Set ws = PanelSheet()

    Dim ticker As String: ticker = ResolveTicker(CStr(ws.Range(TI_TICKER_CELL).Value))
    If ticker = "" Then Exit Sub

    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo ProjErr

    Dim records() As TIRecord, recCount As Long
    Dim lots() As TILot, lotCount As Long
    Call BuildFIFOHistory(ticker, records, recCount, lots, lotCount)
    Call DrawProjection(ws, ticker, lots, lotCount)

ProjDone:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScr
    Exit Sub

ProjErr:
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScr
    If Application.Visible Then
        MsgBox "RefreshTickerProjection error" & vbCrLf & vbCrLf & _
               "Err " & errN & ": " & errD, vbCritical, "Ticker Insight"
    End If
End Sub

' ================================================================
' SHEET ACCESS
' ================================================================
Private Function PanelSheet() As Worksheet
    Set PanelSheet = ThisWorkbook.Sheets(SH_RR4)
End Function

' ----------------------------------------------------------------
' Draws every STATIC element of the panel (labels, section titles,
' history table header, input cells). Values are filled by the
' Draw* sections below.
' ----------------------------------------------------------------
Private Sub DrawShellHeaders(ws As Worksheet, ticker As String)
    ' --- Row 1/2: TICKER <GO> label ABOVE its input (v2.5), company name beside the input ---
    With ws.cells(RR4_TOP + 1, TI_LBL)
        .Value = "TICKER <GO>"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .ShrinkToFit = False
    End With
    With ws.Range(TI_TICKER_CELL)
        .NumberFormat = "@"
        .Value = ticker
        .Interior.Color = RR4_INPUT_BG          ' dark grey = input cell
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
    End With
    If ticker <> "" Then
        With ws.cells(RR4_TOP + 2, TI_LBL + 1)
            .Value = PanelCompanyName(ws, ticker)
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlLeft
        End With
    End If

    ' --- section titles ---
    Call WriteTitle(ws, RR4_TOP + TI_BODY + 3, TI_LBL, "1) ACTIVE POSITION")
    Call WriteTitle(ws, RR4_TOP + TI_BODY + 10, TI_LBL, "2) LIFETIME METRICS")
    Call WriteTitle(ws, RR4_TOP + TI_BODY + 15, TI_LBL, "3) PROJECTION")
    Call WriteTitle(ws, RR4_TOP + TI_BODY + 3, TI_HCOL, "4) RECENT TRADE HISTORY")

    ' --- metric labels (cyan); values go in column K ---
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 4, TI_LBL, "NET EXPOSURE")
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 5, TI_LBL, "ENTRY PRICE (W.AP)")
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 6, TI_LBL, "LAST PRICE")
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 7, TI_LBL, "UNREALIZED PNL")
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 8, TI_LBL, "RETURN %")

    Call WriteLabel(ws, RR4_TOP + TI_BODY + 11, TI_LBL, "REALIZED PNL")
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 12, TI_LBL, "LIFETIME EFF")
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 13, TI_LBL, "VELOCITY/DAY")

    Call WriteLabel(ws, RR4_TOP + TI_BODY + 16, TI_LBL, "LAST PRICE")
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 17, TI_LBL, "PRICE TARGET")
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 18, TI_LBL, "PROJECTED PNL")
    Call WriteLabel(ws, RR4_TOP + TI_BODY + 19, TI_LBL, "PROJECTED RET %")

    ' --- trade-history table header ---
    Dim hdrs As Variant
    hdrs = Array("ENTRY DATE", "EXIT DATE", "DAYS", "SHARES", "REALIZED PNL", "RETURN %")
    Dim c As Long
    For c = 0 To UBound(hdrs)
        With ws.cells(TI_HHDR, TI_HCOL + c)
            .Value = hdrs(c)
            .Font.Color = RGB(150, 150, 150)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
    Next c
    With ws.Range(ws.cells(TI_HHDR, TI_HCOL), ws.cells(TI_HHDR, TI_RIGHT)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    ' --- grey PRICE TARGET input ---
    With ws.Range(TI_TARGET_CELL)
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
    End With
End Sub

' Name from the position log when the ticker is held (no network call),
' otherwise Yahoo via Attach.GetCompanyName.
Private Function PanelCompanyName(ws As Worksheet, ticker As String) As String
    Dim r As Long, normT As String: normT = NormalizeTicker(ticker)
    For r = RR4_POS_FIRST To RR4_POS_FIRST + 300
        Dim t As String: t = CStr(ws.cells(r, RR4_LEFT + 1).Value)
        If t = "" Then Exit For
        If NormalizeTicker(t) = normT Then
            PanelCompanyName = CStr(ws.cells(r, RR4_LEFT + 2).Value)
            Exit Function
        End If
    Next r
    Dim tmp As String: tmp = ticker
    On Error Resume Next
    PanelCompanyName = GetCompanyName(tmp)
    On Error GoTo 0
End Function

' Wipe the panel block only (J5:S25 minus the J6 input cell).
Private Sub ClearTickerData(ws As Worksheet)
    Dim blk As Range
    Set blk = Union(ws.Range(ws.cells(RR4_TOP + 1, TI_LBL), ws.cells(RR4_TOP + 1, TI_RIGHT)), _
                    ws.Range(ws.cells(RR4_TOP + 2, TI_LBL + 1), ws.cells(RR4_TOP + 2, TI_RIGHT)), _
                    ws.Range(ws.cells(RR4_TOP + 3, TI_LBL), ws.cells(TI_BOTTOM, TI_RIGHT)))   ' from row 7: the blank row above the body too
    With blk
        .ClearContents
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Bold = False
        .Font.Italic = False
        .Font.Size = 10
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
        .ShrinkToFit = False
        .Borders.LineStyle = xlNone
        .FormatConditions.Delete
    End With
End Sub

' ================================================================
' FIFO HISTORY BUILDER
' Walks Transactions chronologically per (ticker, broker).
' BUY pushes a lot. SELL pops lots from the front and emits one
' TIRecord per lot consumed (so one SELL that eats 3 lots = 3 rows).
' ADJUSTCOST distributes per share across all current lots.
' ================================================================
Private Sub BuildFIFOHistory(ByVal ticker As String, _
                             ByRef records() As TIRecord, _
                             ByRef recCount As Long, _
                             ByRef lots() As TILot, _
                             ByRef lotCount As Long)
    recCount = 0
    lotCount = 0
    ReDim records(1 To 256)
    ReDim lots(1 To 64)

    Dim wsTr As Worksheet
    On Error Resume Next: Set wsTr = ThisWorkbook.Sheets(SH_TR): On Error GoTo 0
    If wsTr Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = TrLastRow(wsTr)
    If lastRow <= TrHdrRow(wsTr) Then Exit Sub

    ' Collect rows for this ticker (normalised so 2330 / 2330.TW / 2330.TWO match)
    Dim tickerNorm As String: tickerNorm = NormalizeTicker(ticker)
    Dim tRows() As Long, tCount As Long
    ReDim tRows(1 To lastRow)
    Dim r As Long
    For r = TrHdrRow(wsTr) + 1 To lastRow
        If NormalizeTicker(CStr(wsTr.cells(r, TrCol(wsTr, COL_TICKER)).Value)) = tickerNorm Then
            tCount = tCount + 1
            tRows(tCount) = r
        End If
    Next r
    If tCount = 0 Then Exit Sub

    ' Sort by date ascending (insertion sort - tCount usually small)
    Dim i As Long, j As Long, tmp As Long
    For i = 1 To tCount - 1
        For j = i + 1 To tCount
            If SafeDate(wsTr.cells(tRows(i), TrCol(wsTr, COL_DATE)).Value) > _
               SafeDate(wsTr.cells(tRows(j), TrCol(wsTr, COL_DATE)).Value) Then
                tmp = tRows(i): tRows(i) = tRows(j): tRows(j) = tmp
            End If
        Next j
    Next i

    ' Per-broker FIFO lot queue
    Dim queues As Object: Set queues = CreateObject("Scripting.Dictionary")

    Dim k As Long
    For k = 1 To tCount
        r = tRows(k)
        Dim tDate As Date: tDate = SafeDate(wsTr.cells(r, TrCol(wsTr, COL_DATE)).Value)
        Dim Action As String: Action = UCase(Trim(CStr(wsTr.cells(r, TrCol(wsTr, COL_ACTION)).Value)))
        Dim shares As Double: shares = SafeNum(wsTr.cells(r, TrCol(wsTr, COL_SHARES)).Value)
        Dim netAmt As Double: netAmt = SafeNum(wsTr.cells(r, TrCol(wsTr, COL_NETAMT)).Value)
        Dim broker As String: broker = Trim(CStr(wsTr.cells(r, TrCol(wsTr, COL_BROKER)).Value))
        If broker = "" Then broker = "Default"

        If shares <= 0 And Action <> "ADJUSTCOST" Then GoTo NextTx

        If Not queues.Exists(broker) Then queues.Add broker, New Collection
        Dim queue As Collection: Set queue = queues(broker)

        Select Case Action
            Case "BUY"
                Dim cps As Double
                If shares > 0 Then cps = netAmt / shares Else cps = 0
                queue.Add Array(tDate, shares, cps)

            Case "SELL"
                Dim sellPx As Double
                If shares > 0 Then sellPx = netAmt / shares Else sellPx = 0
                Dim remaining As Double: remaining = shares
                Do While remaining > 0.0000001 And queue.count > 0
                    Dim frontLot As Variant: frontLot = queue(1)
                    Dim lotDate As Date: lotDate = CDate(frontLot(0))
                    Dim lotShares As Double: lotShares = CDbl(frontLot(1))
                    Dim lotCps As Double: lotCps = CDbl(frontLot(2))

                    Dim consumed As Double
                    If lotShares <= remaining + 0.0000001 Then
                        consumed = lotShares
                    Else
                        consumed = remaining
                    End If

                    ' Emit a FIFO trade record
                    recCount = recCount + 1
                    If recCount > UBound(records) Then _
                        ReDim Preserve records(1 To UBound(records) * 2)
                    With records(recCount)
                        .EntryDate = lotDate
                        .ExitDate = tDate
                        .DaysHeld = CLng(tDate - lotDate)
                        .shares = consumed
                        .Cost = consumed * lotCps
                        .Proceeds = consumed * sellPx
                        .PnL = .Proceeds - .Cost
                        If .Cost <> 0 Then
                            .ReturnPct = .PnL / .Cost
                        Else
                            .ReturnPct = 0
                        End If
                        .broker = broker
                    End With

                    ' Update / pop the lot
                    If consumed >= lotShares - 0.0000001 Then
                        queue.Remove 1
                    Else
                        queue.Add Array(lotDate, lotShares - consumed, lotCps), Before:=1
                        queue.Remove 2
                    End If
                    remaining = remaining - consumed
                Loop

            Case "ADJUSTCOST"
                Dim totSh As Double: totSh = 0
                Dim li As Long
                For li = 1 To queue.count
                    totSh = totSh + CDbl(queue(li)(1))
                Next li
                If totSh > 0 Then
                    Dim adjPerSh As Double: adjPerSh = netAmt / totSh
                    For li = 1 To queue.count
                        Dim adjLot As Variant: adjLot = queue(li)
                        Dim newCps As Double: newCps = CDbl(adjLot(2)) + adjPerSh
                        queue.Remove li
                        If li > queue.count Then
                            queue.Add Array(adjLot(0), adjLot(1), newCps)
                        Else
                            queue.Add Array(adjLot(0), adjLot(1), newCps), Before:=li
                        End If
                    Next li
                End If
        End Select
NextTx:
    Next k

    If recCount > 0 Then ReDim Preserve records(1 To recCount)

    ' Flatten residual lots (active position)
    Dim brokerKey As Variant
    For Each brokerKey In queues.keys
        Dim bq As Collection: Set bq = queues(brokerKey)
        Dim li2 As Long
        For li2 = 1 To bq.count
            lotCount = lotCount + 1
            If lotCount > UBound(lots) Then _
                ReDim Preserve lots(1 To UBound(lots) * 2)
            With lots(lotCount)
                .lotDate = CDate(bq(li2)(0))
                .shares = CDbl(bq(li2)(1))
                .costPerShare = CDbl(bq(li2)(2))
                .broker = CStr(brokerKey)
            End With
        Next li2
    Next brokerKey

    If lotCount > 0 Then ReDim Preserve lots(1 To lotCount)
End Sub

' ================================================================
' SECTION 1: ACTIVE POSITION (rows 4-8)
' ================================================================
Private Sub DrawActivePosition(ws As Worksheet, ticker As String, _
                                lots() As TILot, lotCount As Long)
    Dim totalShares As Double, totalCost As Double
    Dim i As Long
    For i = 1 To lotCount
        totalShares = totalShares + lots(i).shares
        totalCost = totalCost + lots(i).shares * lots(i).costPerShare
    Next i

    Dim fx As Double: fx = GetFXToTWD(ticker)                   ' RR4!B2 for US, 1 for TW
    Dim lastPx As Double: lastPx = GetStockPriceSafe(ticker)
    Dim wap As Double, unrl As Double, retPct As Double, netExp As Double

    If totalShares > 0 Then
        wap = totalCost / totalShares
        netExp = lastPx * totalShares * fx              ' aggregate -> TWD
        unrl = (lastPx - wap) * totalShares * fx        ' aggregate -> TWD
        If totalCost > 0 Then retPct = (lastPx - wap) * totalShares / totalCost
    End If

    ' NOTE: labels are drawn by DrawShellHeaders. We only fill values here.
    If totalShares > 0 Then
        Call WriteValueMoney(ws, RR4_TOP + TI_BODY + 4, TI_VAL, netExp, "NT$")
        Call WriteValuePrice(ws, RR4_TOP + TI_BODY + 5, TI_VAL, wap)           ' per-share native (no fx)
    Else
        With ws.cells(RR4_TOP + TI_BODY + 4, TI_VAL)
            .Value = "NO ACTIVE POSITION"
            .Font.Color = RGB(150, 150, 150)
            .Font.Italic = True
            .Interior.Color = RGB(0, 0, 0)
            .Borders.LineStyle = xlNone
            .ShrinkToFit = False
        End With
        Call WriteDash(ws, RR4_TOP + TI_BODY + 5, TI_VAL)
    End If

    If totalShares > 0 And lastPx > 0 Then
        Call WriteValuePrice(ws, RR4_TOP + TI_BODY + 6, TI_VAL, lastPx)        ' per-share native (no fx)
        Call WriteValuePnL(ws, RR4_TOP + TI_BODY + 7, TI_VAL, unrl, "NT$")
        Call WriteValuePct(ws, RR4_TOP + TI_BODY + 8, TI_VAL, retPct)
    Else
        Call WriteDash(ws, RR4_TOP + TI_BODY + 6, TI_VAL)
        Call WriteDash(ws, RR4_TOP + TI_BODY + 7, TI_VAL)
        Call WriteDash(ws, RR4_TOP + TI_BODY + 8, TI_VAL)
    End If
End Sub

' ================================================================
' SECTION 2: LIFETIME METRICS (rows 11-13)
' LIFETIME EFF = Profit Factor (gross win / gross loss)
' VELOCITY/DAY = Realized PnL / sum(holding days), min 1 per trade
' ================================================================
Private Sub DrawLifetimeMetrics(ws As Worksheet, ticker As String, _
                                 records() As TIRecord, recCount As Long)
    Dim totalPnL As Double, grossWin As Double, grossLoss As Double
    Dim totalDays As Long
    Dim i As Long
    For i = 1 To recCount
        totalPnL = totalPnL + records(i).PnL
        If records(i).PnL > 0 Then grossWin = grossWin + records(i).PnL
        If records(i).PnL < 0 Then grossLoss = grossLoss + Abs(records(i).PnL)
        Dim d As Long: d = records(i).DaysHeld
        If d < 1 Then d = 1
        totalDays = totalDays + d
    Next i

    Dim fx As Double: fx = GetFXToTWD(ticker)

    ' --- REALIZED PNL (TWD-converted) ---
    Call WriteValuePnL(ws, RR4_TOP + TI_BODY + 11, TI_VAL, totalPnL * fx, "NT$")

    ' --- LIFETIME EFF (Profit Factor) ---
    ' Defensive: clear borders/background, set NumberFormat BEFORE Value,
    ' and Round the value so even if format application fails the cell
    ' still displays 2 decimals max.
    With ws.cells(RR4_TOP + TI_BODY + 12, TI_VAL)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft
        If grossLoss > 0 Then
            Dim pf As Double: pf = grossWin / grossLoss
            .NumberFormat = "0.00"
            .Value = Round(pf, 2)
            If pf >= 1 Then
                .Font.Color = RGB(0, 210, 100)
            Else
                .Font.Color = RGB(255, 80, 80)
            End If
            .Font.Bold = True
        ElseIf grossWin > 0 Then
            .NumberFormat = "@"
            .Value = "infinite"     ' only wins, no losses
            .Font.Color = RGB(0, 210, 100)
            .Font.Bold = True
        Else
            .NumberFormat = "@"
            .Value = ChrW(&H2014)   ' em dash (consistent with WriteDash)
            .Font.Color = RGB(150, 150, 150)
        End If
    End With

    ' --- VELOCITY/DAY (TWD-converted, PnL per holding day) ---
    Dim velocity As Double
    If totalDays > 0 Then velocity = totalPnL / totalDays
    Call WriteValuePnL(ws, RR4_TOP + TI_BODY + 13, TI_VAL, velocity * fx, "NT$")
End Sub

' ================================================================
' SECTION 3: PROJECTION (rows 16-19)
' K17 is the user-editable PRICE TARGET (grey input background)
' ================================================================
Private Sub DrawProjection(ws As Worksheet, ticker As String, _
                            lots() As TILot, lotCount As Long)
    Dim totalShares As Double
    Dim i As Long
    For i = 1 To lotCount: totalShares = totalShares + lots(i).shares: Next i

    Dim lastPx As Double: lastPx = GetStockPriceSafe(ticker)
    Dim fx As Double: fx = GetFXToTWD(ticker)

    ' LAST PRICE (per-share native)
    If lastPx > 0 Then
        Call WriteValuePrice(ws, RR4_TOP + TI_BODY + 16, TI_VAL, lastPx)
    Else
        Call WriteDash(ws, RR4_TOP + TI_BODY + 16, TI_VAL)
    End If

    ' PRICE TARGET (per-share native input) - preserve existing value
    With ws.Range(TI_TARGET_CELL)
        If Not IsNumeric(.Value) Then .Value = vbNullString
    End With

    Dim tgt As Double
    If IsNumeric(ws.Range(TI_TARGET_CELL).Value) Then tgt = CDbl(ws.Range(TI_TARGET_CELL).Value)

    ' PROJECTED PNL (TWD-converted)
    If totalShares > 0 And tgt > 0 And lastPx > 0 Then
        Dim projPnL As Double: projPnL = (tgt - lastPx) * totalShares * fx
        Call WriteValuePnL(ws, RR4_TOP + TI_BODY + 18, TI_VAL, projPnL, "NT$")
    Else
        Call WriteDash(ws, RR4_TOP + TI_BODY + 18, TI_VAL)
    End If

    ' PROJECTED RETURN %
    If tgt > 0 And lastPx > 0 Then
        Dim projRet As Double: projRet = (tgt / lastPx) - 1
        Call WriteValuePct(ws, RR4_TOP + TI_BODY + 19, TI_VAL, projRet)
    Else
        Call WriteDash(ws, RR4_TOP + TI_BODY + 19, TI_VAL)
    End If
End Sub

' ================================================================
' SECTION 4: RECENT TRADE HISTORY (M5:R22, newest first)
' ================================================================
Private Sub DrawTradeHistory(ws As Worksheet, ticker As String, _
                              records() As TIRecord, recCount As Long)
    If recCount = 0 Then Exit Sub

    ' Sort by ExitDate descending (newest first)
    Dim i As Long, j As Long, tmp As TIRecord
    For i = 1 To recCount - 1
        For j = i + 1 To recCount
            If records(j).ExitDate > records(i).ExitDate Then
                tmp = records(i): records(i) = records(j): records(j) = tmp
            End If
        Next j
    Next i

    Dim fx As Double: fx = GetFXToTWD(ticker)
    Dim pnlFmtStr As String: pnlFmtStr = PnLFmt("NT$")   ' quoted-literal NT$ format

    Dim r As Long: r = TI_H1
    Dim c As Long: c = TI_HCOL
    For i = 1 To recCount
        If r > TI_HLAST Then Exit For
        With ws.cells(r, c)
            .NumberFormat = "yyyy-mm-dd"
            .Value = records(i).EntryDate
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlLeft
        End With
        With ws.cells(r, c + 1)
            .NumberFormat = "yyyy-mm-dd"
            .Value = records(i).ExitDate
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlLeft
        End With
        With ws.cells(r, c + 2)
            .Value = records(i).DaysHeld
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlCenter
        End With
        With ws.cells(r, c + 3)
            .NumberFormat = "#,##0.##"
            .Value = records(i).shares
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlRight
        End With
        ' --- PnL converted to TWD ---
        Dim pnlTWD As Double: pnlTWD = records(i).PnL * fx
        With ws.cells(r, c + 4)
            .NumberFormat = pnlFmtStr
            .Value = pnlTWD
            .Font.Color = USPnLColor(pnlTWD)
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .ShrinkToFit = True
        End With
        With ws.cells(r, c + 5)
            .NumberFormat = "0.00%"
            .Value = Round(records(i).ReturnPct, 4)
            .Font.Color = USPnLColor(records(i).PnL)
            .HorizontalAlignment = xlRight
        End With
        r = r + 1
    Next i

    If recCount > TI_HLAST - TI_H1 + 1 Then
        With ws.cells(TI_BOTTOM, c)
            .Value = "+" & (recCount - (TI_HLAST - TI_H1 + 1)) & " older trades not shown"
            .Font.Color = RGB(150, 150, 150)
            .Font.Italic = True
            .Font.Size = 9
        End With
    End If
End Sub

' ================================================================
' CELL HELPERS
' ================================================================
Private Sub WriteTitle(ws As Worksheet, r As Long, c As Long, txt As String)
    With ws.cells(r, c)
        .Value = txt
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
    End With
End Sub

Private Sub WriteLabel(ws As Worksheet, r As Long, c As Long, txt As String)
    With ws.cells(r, c)
        .Value = txt
        .Font.Color = RGB(0, 200, 255)
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
    End With
End Sub

' All value-writers below set NumberFormat BEFORE Value and explicitly
' clear borders/background to prevent stale formatting from prior renders.
' Column K is narrow (it is also the position log's UNRL PNL column), so
' values shrink to fit instead of turning into ####.
'
' Currency symbol "NT$" must be quoted in the format string. Bare
' "$" in Excel format codes is treated as a locale-aware currency
' specifier (in zh-TW it expands to "NT$"), so an unquoted "NT$..."
' was being mis-parsed and silently rejected on some Excel versions.
' Using Chr(34) (double quote) around the symbol forces literal text.
Private Sub WriteValueMoney(ws As Worksheet, r As Long, c As Long, _
                             v As Double, sym As String)
    With ws.cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = MoneyFmt(sym)
        .Value = Round(v, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .ShrinkToFit = True
    End With
End Sub

Private Sub WriteValuePrice(ws As Worksheet, r As Long, c As Long, v As Double)
    With ws.cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = "#,##0.00"
        .Value = Round(v, 2)
        .Font.Color = RGB(221, 221, 221)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .ShrinkToFit = True
    End With
End Sub

Private Sub WriteValuePnL(ws As Worksheet, r As Long, c As Long, _
                           v As Double, sym As String)
    With ws.cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = PnLFmt(sym)
        .Value = Round(v, 0)
        .Font.Color = USPnLColor(v)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .ShrinkToFit = True
    End With
End Sub

' Quoted-literal currency format builders.
' For sym="NT$", returns:  "NT$"#,##0
' That outer quote pair is part of the Excel format code itself.
Private Function MoneyFmt(sym As String) As String
    Dim q As String: q = Chr(34) & sym & Chr(34)
    MoneyFmt = q & "#,##0"
End Function

Private Function PnLFmt(sym As String) As String
    Dim q As String: q = Chr(34) & sym & Chr(34)
    PnLFmt = q & "#,##0;-" & q & "#,##0"
End Function

Private Sub WriteValuePct(ws As Worksheet, r As Long, c As Long, v As Double)
    With ws.cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = "0.00%"
        .Value = Round(v, 4)
        .Font.Color = USPnLColor(v)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .ShrinkToFit = True
    End With
End Sub

Private Sub WriteDash(ws As Worksheet, r As Long, c As Long)
    With ws.cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = "@"
        .Value = ChrW(&H2014)   ' em dash
        .Font.Color = RGB(150, 150, 150)
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
    End With
End Sub

' US/Bloomberg color convention: green positive, red negative
Private Function USPnLColor(v As Double) As Long
    If v > 0 Then
        USPnLColor = RGB(0, 210, 100)
    ElseIf v < 0 Then
        USPnLColor = RGB(255, 80, 80)
    Else
        USPnLColor = RGB(200, 200, 200)
    End If
End Function

' ================================================================
' TICKER / CURRENCY HELPERS
' ================================================================

' Strip .TW / .TWO suffix and uppercase so "2330", "2330.TW", "2330.TWO"
' all normalise to the same key. Used by FIFO matching and ticker classification.
Private Function NormalizeTicker(t As String) As String
    Dim n As String: n = UCase(Trim(CStr(t)))
    If Len(n) >= 4 Then
        If Right(n, 4) = ".TWO" Then n = Left(n, Len(n) - 4)
    End If
    If Len(n) >= 3 Then
        If Right(n, 3) = ".TW" Then n = Left(n, Len(n) - 3)
    End If
    NormalizeTicker = n
End Function

' Canonical ticker for a typed value (v2.4, 2026-09-13): the market is
' decided by what the Transactions log actually holds, not by how the
' code was typed. "3374" typed bare matched its 3374.TWO rows through
' NormalizeTicker, but IsTWTicker saw no suffix and every aggregate got
' multiplied by USD/TWD (816 -> 25,808). So: a bare code takes the suffix
' of its first Transactions row; with no rows at all, an all-digit code
' (optionally one trailing letter, e.g. 00981A) is assumed TW (.TW).
Private Function ResolveTicker(ByVal ticker As String) As String
    Dim t As String: t = UCase(Trim(CStr(ticker)))
    ResolveTicker = t
    If t = "" Then Exit Function
    If InStr(t, ".TW") > 0 Then Exit Function
    Dim wsTr As Worksheet
    On Error Resume Next: Set wsTr = ThisWorkbook.Sheets(SH_TR): On Error GoTo 0
    If Not wsTr Is Nothing Then
        Dim lastRow As Long: lastRow = TrLastRow(wsTr)
        Dim r As Long, raw As String
        For r = TrHdrRow(wsTr) + 1 To lastRow
            raw = UCase(Trim(CStr(wsTr.cells(r, TrCol(wsTr, COL_TICKER)).Value)))
            If NormalizeTicker(raw) = t Then
                If InStr(raw, ".TW") > 0 Then ResolveTicker = raw
                Exit Function
            End If
        Next r
    End If
    ' no log rows: TW codes are digits (+ optional one letter), US codes are letters
    Dim core As String: core = t
    If Len(core) > 1 Then
        If Right(core, 1) Like "[A-Z]" Then core = Left(core, Len(core) - 1)
    End If
    If Len(core) >= 4 And core Like String(Len(core), "#") Then ResolveTicker = t & ".TW"
End Function

' Suffix-based classification:
'   .TW or .TWO suffix -> TW (no FX conversion)
'   Everything else    -> US (FX applied to aggregates)
'
' Tickers reaching this function have been through ResolveTicker, so a
' bare TW code already carries the suffix its Transactions rows use.
Private Function IsTWTicker(ticker As String) As Boolean
    Dim raw As String: raw = UCase(Trim(CStr(ticker)))
    IsTWTicker = (InStr(raw, ".TW") > 0)
End Function

' Per-share PRICE currency symbol (native; never FX-converted)
'   TW -> "NT$"     US -> "$"
Private Function GetPriceSymbol(ticker As String) As String
    If IsTWTicker(ticker) Then GetPriceSymbol = "NT$" Else GetPriceSymbol = "$"
End Function

' Multiplier for aggregate PnL / exposure -> TWD
'   TW ticker -> 1
'   US ticker -> RR4!B2 (the rate the dashboard uses), else FX_USD_TWD
Private Function GetFXToTWD(ticker As String) As Double
    If IsTWTicker(ticker) Then
        GetFXToTWD = 1
        Exit Function
    End If
    Dim fx As Double
    On Error Resume Next
    fx = RR4FxRate()            ' the page's USD/TWD input, wherever it sits
    On Error GoTo 0
    If fx < 20 Or fx > 50 Then fx = FX_USD_TWD
    GetFXToTWD = fx
End Function

Private Function GetStockPriceSafe(ticker As String) As Double
    GetStockPriceSafe = 0
    On Error Resume Next
    Dim raw As String: raw = UCase(Trim(CStr(ticker)))
    If InStr(raw, ".TW") > 0 Then
        ' .TW / .TWO -> TWSE / TPEx OpenAPI (TaiwanPriceFetcher module)
        GetStockPriceSafe = GetTWStockPrice(ticker)
    Else
        ' US stocks / ETFs / indices -> Yahoo Finance (Attach module)
        GetStockPriceSafe = GetStockPrice(ticker)
    End If
    On Error GoTo 0
End Function

' Safer date / number conversion (returns 0 if not parseable)
Private Function SafeDate(v As Variant) As Date
    On Error Resume Next
    If IsDate(v) Then
        SafeDate = CDate(v)
    ElseIf IsNumeric(v) Then
        SafeDate = CDate(CDbl(v))
    End If
    On Error GoTo 0
End Function

Private Function SafeNum(v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then SafeNum = CDbl(v)
    On Error GoTo 0
End Function
