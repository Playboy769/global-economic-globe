Attribute VB_Name = "TickerInsight"
Option Explicit

' ================================================================
' TICKER INSIGHT v1.0
' Per-ticker FIFO trade history & performance dashboard
' Bloomberg terminal style, US color convention (green +, red -)
'
' ENTRY POINTS:
'   Sub OpenTickerInsight()          - call from Alt+F8 or button
'   Sub RefreshTickerInsight(ticker) - full rebuild for a ticker
'   Sub RefreshTickerProjection      - recompute PROJECTION block only
'   Sub TickerInsight_OnChange(t)    - dispatcher for sheet events
'
' WORKSHEET LAYOUT (auto-created on first run):
'   Row  1: TICKER <GO>: [B1 input]
'   Row  3: 3 section titles  (A=ACTIVE | D=LIFETIME | G=PROJECTION)
'   Row  4: NET EXPOSURE   | REALIZED PNL  | LAST PRICE
'   Row  5: ENTRY (W.AP)   | LIFETIME EFF  | PRICE TARGET [H5 input]
'   Row  6: LAST PRICE     | VELOCITY/DAY  | PROJECTED PNL
'   Row  7: UNREALIZED PNL |               | PROJECTED RETURN %
'   Row  8: RETURN %       |               |
'   Row 10: 4) RECENT TRADE HISTORY
'   Row 11: ENTRY | EXIT | DAYS | SHARES | PNL | RETURN % | BROKER
'   Row 12+: ... newest-first ...
'
' INSTALLATION:
'   1. VBE > File > Import File > TickerInsight.bas (this file)
'   2. Run OpenTickerInsight (Alt+F8) - "Ticker Insight" sheet created
'   3. In VBE left panel, double-click the "Ticker Insight" sheet
'      under "Microsoft Excel Objects". Paste this Worksheet_Change:
'
'        Private Sub Worksheet_Change(ByVal Target As Range)
'            Dim hitB1 As Boolean
'            Dim hitH5 As Boolean
'            hitB1 = Not Intersect(Target, Me.Range("B1")) Is Nothing
'            hitH5 = Not Intersect(Target, Me.Range("H5")) Is Nothing
'            If Not (hitB1 Or hitH5) Then Exit Sub
'            If Target.Count > 1 Then Exit Sub
'            If hitB1 And Trim(Me.Range("B1").Value) = "" Then Exit Sub
'
'            Application.EnableEvents = False
'            On Error GoTo ErrorHandler
'
'            If hitB1 Then
'                Call RefreshTickerInsight(CStr(Me.Range("B1").Value))
'            ElseIf hitH5 Then
'                Call RefreshTickerProjection
'            End If
'
'        ErrorHandler:
'            Application.EnableEvents = True
'            If Err.Number <> 0 Then
'                MsgBox "Ticker Insight error: " & Err.Description, _
'                       vbCritical, "Err " & Err.Number
'            End If
'        End Sub
'
'   4. Type a ticker in B1 (e.g. CRDO, 2330) and optionally H5 target.
' ================================================================

Private Const SH_TI   As String = "Ticker Insight"
Private Const SH_TR   As String = "Transactions"
Private Const SH_REAL As String = "Realized"
Private Const SH_RR4  As String = "RR4"

' Transaction column letters (from existing schema)
Private Const COL_DATE   As String = "B"
Private Const COL_TICKER As String = "C"
Private Const COL_ACTION As String = "D"
Private Const COL_SHARES As String = "E"
Private Const COL_NETAMT As String = "I"
Private Const COL_BROKER As String = "N"

' Exchange rate: USD -> TWD. All aggregate PnL / exposure values for
' US tickers are multiplied by this constant before display. Per-share
' prices (Entry, Last, Price Target) stay in NATIVE currency.
Private Const FX_USD_TWD As Double = 31.6

' ----------------------------------------------------------------
' One FIFO-paired trade record (one BUY lot matched to one SELL)
' ----------------------------------------------------------------
Private Type TIRecord
    EntryDate As Date
    ExitDate  As Date
    DaysHeld  As Long
    Shares    As Double
    Cost      As Double
    Proceeds  As Double
    PnL       As Double
    ReturnPct As Double
    Broker    As String
End Type

' ----------------------------------------------------------------
' One residual lot still open after all transactions processed
' ----------------------------------------------------------------
Private Type TILot
    LotDate      As Date
    Shares       As Double
    CostPerShare As Double
    Broker       As String
End Type

' ================================================================
' PUBLIC ENTRY: open the dashboard (auto-creates sheet)
' Event hooking is done via Worksheet_Change pasted on the sheet's
' code page (see install snippet in module header).
' ================================================================
Public Sub OpenTickerInsight()
    Dim ws As Worksheet
    Set ws = EnsureTickerInsightSheet()

    ' One-shot defensive cleanup: nuke any stray shapes/conditional
    ' formats inherited from previous versions of the sheet.
    Application.EnableEvents = False
    Call RemoveSheetShapes(ws)
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    ws.AutoFilterMode = False
    On Error GoTo 0
    Application.EnableEvents = True

    ws.Activate
    ws.Range("B1").Select

    ' If sheet already has a ticker, do a full refresh.
    ' Otherwise just redraw the empty shell so the user always sees structure.
    Dim existing As String: existing = Trim(CStr(ws.Range("B1").Value))
    If existing <> "" Then
        Call RefreshTickerInsight(existing)
    Else
        Application.EnableEvents = False
        Call DrawShellHeaders(ws, vbNullString)
        Application.EnableEvents = True
    End If
End Sub

' ================================================================
' DIAGNOSTIC: dumps the state of the Ticker Insight pipeline for the
' ticker currently in B1 (or a passed-in one). Useful when no values
' appear in the dashboard. Run via Alt+F8 -> DiagnoseTicker.
' ================================================================
Public Sub DiagnoseTicker(Optional ByVal forcedTicker As String = "")
    Dim msg As String
    Dim lf As String: lf = vbCrLf

    On Error GoTo ErrTrap

    ' --- 1. Sheets ---
    msg = msg & "=== SHEET CHECK ===" & lf
    Dim wsTI As Worksheet, wsTr As Worksheet
    On Error Resume Next
    Set wsTI = ThisWorkbook.Sheets(SH_TI)
    Set wsTr = ThisWorkbook.Sheets(SH_TR)
    On Error GoTo ErrTrap

    msg = msg & "Ticker Insight: " & IIf(wsTI Is Nothing, "MISSING", "OK") & lf
    msg = msg & "Transactions  : " & IIf(wsTr Is Nothing, "MISSING", "OK") & lf

    If wsTr Is Nothing Then GoTo ShowMsg

    ' --- 2. Ticker resolution ---
    msg = msg & lf & "=== TICKER ===" & lf
    Dim ticker As String
    If forcedTicker <> "" Then
        ticker = forcedTicker
    ElseIf Not (wsTI Is Nothing) Then
        ticker = CStr(wsTI.Range("B1").Value)
    End If
    ticker = UCase(Trim(ticker))

    If ticker = "" Then
        msg = msg & "(B1 empty - type a ticker then re-run, or DiagnoseTicker " & _
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
    lastRow = wsTr.Cells(wsTr.Rows.Count, "A").End(xlUp).Row
    msg = msg & "Total rows    : " & (lastRow - 1) & lf

    Dim normT As String: normT = NormalizeTicker(ticker)
    Dim matchCount As Long, sample As String, samples As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim cellT As String: cellT = CStr(wsTr.Cells(r, COL_TICKER).Value)
        If NormalizeTicker(cellT) = normT Then
            matchCount = matchCount + 1
            If samples < 5 Then
                sample = sample & "  R" & r & ":  " & _
                    CStr(wsTr.Cells(r, COL_DATE).Value) & "  " & _
                    "[" & cellT & "]  " & _
                    CStr(wsTr.Cells(r, COL_ACTION).Value) & "  " & _
                    "sh=" & CStr(wsTr.Cells(r, COL_SHARES).Value) & "  " & _
                    "net=" & CStr(wsTr.Cells(r, COL_NETAMT).Value) & "  " & _
                    "brk=" & CStr(wsTr.Cells(r, COL_BROKER).Value) & lf
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
        For r = 2 To lastRow
            If cnt < 10 Then
                Dim t2 As String: t2 = CStr(wsTr.Cells(r, COL_TICKER).Value)
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
        msg = msg & "  Days=" & records(1).DaysHeld & "  Shares=" & records(1).Shares & lf
        msg = msg & "  Cost=" & records(1).Cost & "  Proceeds=" & records(1).Proceeds & lf
        msg = msg & "  PnL=" & records(1).PnL & "  Ret%=" & records(1).ReturnPct & lf
        msg = msg & "  Broker=" & records(1).Broker & lf
    End If

    If lotCount > 0 Then
        msg = msg & lf & "First open lot:" & lf
        msg = msg & "  Date=" & lots(1).LotDate & "  Shares=" & lots(1).Shares & lf
        msg = msg & "  CPS=" & lots(1).CostPerShare & "  Broker=" & lots(1).Broker & lf
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
' DISPATCHER: called from ThisWorkbook.Workbook_SheetChange
' ================================================================
Public Sub TickerInsight_OnChange(ByVal Target As Range)
    On Error Resume Next
    If Target.CountLarge > 1 Then Exit Sub
    Select Case Target.Address
        Case "$B$1"
            Call RefreshTickerInsight(CStr(Target.Value))
        Case "$H$5"
            Call RefreshTickerProjection
    End Select
End Sub

' ================================================================
' FULL REFRESH: rebuild every section for a ticker
' ================================================================
Public Sub RefreshTickerInsight(ByVal ticker As String)
    ticker = UCase(Trim(ticker))
    If ticker = "" Then Exit Sub

    Dim ws As Worksheet: Set ws = EnsureTickerInsightSheet()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo ErrTrap

    ' --- Detect ticker change via hidden A2 tracker ------------
    ' A2 stores the PREVIOUS ticker (invisible: black-on-black).
    ' B1 has already been updated by the user, so we can't use it.
    Dim oldTicker As String
    oldTicker = UCase(Trim(CStr(ws.Cells(2, 1).Value)))
    Dim savedTarget As Variant: savedTarget = ws.Cells(5, 8).Value

    Call ClearTickerData(ws)
    Call DrawShellHeaders(ws, ticker)

    ' If user is refreshing the SAME ticker, preserve PRICE TARGET.
    ' If they switched tickers, leave H5 empty (set by shell).
    If oldTicker = ticker And IsNumeric(savedTarget) Then
        If CDbl(savedTarget) > 0 Then ws.Cells(5, 8).Value = CDbl(savedTarget)
    End If

    Dim records() As TIRecord, recCount As Long
    Dim lots() As TILot, lotCount As Long
    Call BuildFIFOHistory(ticker, records, recCount, lots, lotCount)

    If recCount = 0 And lotCount = 0 Then
        With ws.Cells(4, 1)
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
    ' Stamp current ticker into invisible A2 for next refresh's compare
    With ws.Cells(2, 1)
        .Value = ticker
        .Font.Color = RGB(0, 0, 0)
        .Interior.Color = RGB(0, 0, 0)
    End With

Done:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrTrap:
    ' Surface errors instead of silently swallowing - vital for debug
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "RefreshTickerInsight error" & vbCrLf & vbCrLf & _
           "Err " & errN & ": " & errD & vbCrLf & _
           "Ticker: " & ticker, vbCritical, "Ticker Insight"
End Sub

' ================================================================
' LIGHT REFRESH: recompute PROJECTION only (price target changed)
' ================================================================
Public Sub RefreshTickerProjection()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(SH_TI): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim ticker As String: ticker = UCase(Trim(CStr(ws.Range("B1").Value)))
    If ticker = "" Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo ProjErr

    Dim records() As TIRecord, recCount As Long
    Dim lots() As TILot, lotCount As Long
    Call BuildFIFOHistory(ticker, records, recCount, lots, lotCount)
    Call DrawProjection(ws, ticker, lots, lotCount)

ProjDone:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ProjErr:
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "RefreshTickerProjection error" & vbCrLf & vbCrLf & _
           "Err " & errN & ": " & errD, vbCritical, "Ticker Insight"
End Sub

' ================================================================
' SHEET MANAGEMENT
' ================================================================
Private Function EnsureTickerInsightSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_TI)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SH_TI
        Call InitSheetLayout(ws)
    End If
    Set EnsureTickerInsightSheet = ws
End Function

Private Sub InitSheetLayout(ws As Worksheet)
    ' Suppress events around shell setup so writing B1 doesn't recursively
    ' trigger Workbook_SheetChange while the sheet is half-built.
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 10
    End With

    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    ' Column widths: 3 metric groups + history area
    ws.Columns("A").ColumnWidth = 22   ' Label / Entry date
    ws.Columns("B").ColumnWidth = 16   ' Value / Exit date
    ws.Columns("C").ColumnWidth = 8    ' Days
    ws.Columns("D").ColumnWidth = 22   ' Label / Shares
    ws.Columns("E").ColumnWidth = 16   ' Value / PnL
    ws.Columns("F").ColumnWidth = 10   ' Return %
    ws.Columns("G").ColumnWidth = 22   ' Label / Broker
    ws.Columns("H").ColumnWidth = 16   ' Value

    On Error Resume Next
    ws.Tab.Color = RGB(255, 192, 0)
    On Error GoTo 0

    ' Draw the full static shell immediately so the user sees structure
    ' (even before they type a ticker). RefreshTickerInsight will repaint
    ' this same shell + fill in values.
    Call DrawShellHeaders(ws, vbNullString)

    Application.EnableEvents = prevEvents
End Sub

' ----------------------------------------------------------------
' Draws every STATIC element of the dashboard (labels, section
' titles, history table header). Called from InitSheetLayout (empty
' ticker) and from RefreshTickerInsight (with current ticker).
' ----------------------------------------------------------------
Private Sub DrawShellHeaders(ws As Worksheet, ticker As String)
    ' --- Row 1: TICKER <GO>: input ---
    With ws.Cells(1, 1)
        .Value = "TICKER <GO>:"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 11
    End With
    With ws.Cells(1, 2)
        .Value = ticker
        .Interior.Color = RGB(255, 192, 0)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 22

    ' --- Row 3: section titles ---
    With ws.Cells(3, 1)
        .Value = "1) ACTIVE POSITION"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    With ws.Cells(3, 4)
        .Value = "2) LIFETIME METRICS"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    With ws.Cells(3, 7)
        .Value = "3) PROJECTION"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ws.Rows(3).RowHeight = 20

    ' --- Row 4-8: metric labels (cyan) ---
    Call WriteLabel(ws, 4, 1, "NET EXPOSURE")
    Call WriteLabel(ws, 5, 1, "ENTRY PRICE (W.AP)")
    Call WriteLabel(ws, 6, 1, "LAST PRICE")
    Call WriteLabel(ws, 7, 1, "UNREALIZED PNL")
    Call WriteLabel(ws, 8, 1, "RETURN %")

    Call WriteLabel(ws, 4, 4, "REALIZED PNL")
    Call WriteLabel(ws, 5, 4, "LIFETIME EFF")
    Call WriteLabel(ws, 6, 4, "VELOCITY/DAY")

    Call WriteLabel(ws, 4, 7, "LAST PRICE")
    Call WriteLabel(ws, 5, 7, "PRICE TARGET")
    Call WriteLabel(ws, 6, 7, "PROJECTED PNL")
    Call WriteLabel(ws, 7, 7, "PROJECTED RETURN %")

    ' Row heights for metric block
    Dim rh As Long
    For rh = 4 To 8: ws.Rows(rh).RowHeight = 18: Next rh

    ' --- Row 10: history title ---
    With ws.Cells(10, 1)
        .Value = "4) RECENT TRADE HISTORY"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ws.Rows(10).RowHeight = 20

    ' --- Row 11: history table header ---
    Dim hdrs As Variant
    hdrs = Array("ENTRY DATE", "EXIT DATE", "DAYS", "SHARES", _
                 "REALIZED PNL", "RETURN %", "BROKER")
    Dim c As Long
    For c = 0 To UBound(hdrs)
        With ws.Cells(11, c + 1)
            .Value = hdrs(c)
            .Font.Color = RGB(150, 150, 150)
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
        End With
    Next c
    With ws.Range(ws.Cells(11, 1), ws.Cells(11, UBound(hdrs) + 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(255, 192, 0)
        .Weight = xlThin
    End With
    ws.Rows(11).RowHeight = 18

    ' --- Yellow background for H5 (PRICE TARGET input) ---
    With ws.Cells(5, 8)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Size = 11
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub ClearTickerData(ws As Worksheet)
    ' Wipe everything except row 1 (header) and row 2 (gutter)
    With ws.Range(ws.Cells(3, 1), ws.Cells(500, 8))
        .ClearContents
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Bold = False
        .Font.Italic = False
        .Font.Size = 10
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
        .Borders.LineStyle = xlNone
        ' Strip any leftover conditional formatting (icon sets, data bars)
        .FormatConditions.Delete
    End With

    ' Remove ALL shapes (charts, autoshapes, sparklines, smart tags etc.)
    ' that may have been left on the sheet. We own this sheet entirely.
    Call RemoveSheetShapes(ws)

    ' Turn off any AutoFilter that may have been applied
    On Error Resume Next
    ws.AutoFilterMode = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' Deletes every Shape on the worksheet (charts, autoshapes, comment
' callouts, smart tags, sparklines). Loops backwards to avoid index
' shifting as items are removed.
' ----------------------------------------------------------------
Private Sub RemoveSheetShapes(ws As Worksheet)
    On Error Resume Next
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
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
    lastRow = wsTr.Cells(wsTr.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Collect rows for this ticker (normalised so 2330 / 2330.TW / 2330.TWO match)
    Dim tickerNorm As String: tickerNorm = NormalizeTicker(ticker)
    Dim tRows() As Long, tCount As Long
    ReDim tRows(1 To lastRow)
    Dim r As Long
    For r = 2 To lastRow
        If NormalizeTicker(CStr(wsTr.Cells(r, COL_TICKER).Value)) = tickerNorm Then
            tCount = tCount + 1
            tRows(tCount) = r
        End If
    Next r
    If tCount = 0 Then Exit Sub

    ' Sort by date ascending (insertion sort - tCount usually small)
    Dim i As Long, j As Long, tmp As Long
    For i = 1 To tCount - 1
        For j = i + 1 To tCount
            If SafeDate(wsTr.Cells(tRows(i), COL_DATE).Value) > _
               SafeDate(wsTr.Cells(tRows(j), COL_DATE).Value) Then
                tmp = tRows(i): tRows(i) = tRows(j): tRows(j) = tmp
            End If
        Next j
    Next i

    ' Per-broker FIFO lot queue
    Dim queues As Object: Set queues = CreateObject("Scripting.Dictionary")

    Dim k As Long
    For k = 1 To tCount
        r = tRows(k)
        Dim tDate As Date: tDate = SafeDate(wsTr.Cells(r, COL_DATE).Value)
        Dim action As String: action = UCase(Trim(CStr(wsTr.Cells(r, COL_ACTION).Value)))
        Dim shares As Double: shares = SafeNum(wsTr.Cells(r, COL_SHARES).Value)
        Dim netAmt As Double: netAmt = SafeNum(wsTr.Cells(r, COL_NETAMT).Value)
        Dim broker As String: broker = Trim(CStr(wsTr.Cells(r, COL_BROKER).Value))
        If broker = "" Then broker = "Default"

        If shares <= 0 And action <> "ADJUSTCOST" Then GoTo NextTx

        If Not queues.Exists(broker) Then queues.Add broker, New Collection
        Dim queue As Collection: Set queue = queues(broker)

        Select Case action
            Case "BUY"
                Dim cps As Double
                If shares > 0 Then cps = netAmt / shares Else cps = 0
                queue.Add Array(tDate, shares, cps)

            Case "SELL"
                Dim sellPx As Double
                If shares > 0 Then sellPx = netAmt / shares Else sellPx = 0
                Dim remaining As Double: remaining = shares
                Do While remaining > 0.0000001 And queue.Count > 0
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
                        .Shares = consumed
                        .Cost = consumed * lotCps
                        .Proceeds = consumed * sellPx
                        .PnL = .Proceeds - .Cost
                        If .Cost <> 0 Then
                            .ReturnPct = .PnL / .Cost
                        Else
                            .ReturnPct = 0
                        End If
                        .Broker = broker
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
                For li = 1 To queue.Count
                    totSh = totSh + CDbl(queue(li)(1))
                Next li
                If totSh > 0 Then
                    Dim adjPerSh As Double: adjPerSh = netAmt / totSh
                    For li = 1 To queue.Count
                        Dim adjLot As Variant: adjLot = queue(li)
                        Dim newCps As Double: newCps = CDbl(adjLot(2)) + adjPerSh
                        queue.Remove li
                        If li > queue.Count Then
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
    For Each brokerKey In queues.Keys
        Dim bq As Collection: Set bq = queues(brokerKey)
        Dim li2 As Long
        For li2 = 1 To bq.Count
            lotCount = lotCount + 1
            If lotCount > UBound(lots) Then _
                ReDim Preserve lots(1 To UBound(lots) * 2)
            With lots(lotCount)
                .LotDate = CDate(bq(li2)(0))
                .Shares = CDbl(bq(li2)(1))
                .CostPerShare = CDbl(bq(li2)(2))
                .Broker = CStr(brokerKey)
            End With
        Next li2
    Next brokerKey

    If lotCount > 0 Then ReDim Preserve lots(1 To lotCount)
End Sub

' ================================================================
' SECTION 1: ACTIVE POSITION (columns A-B)
' ================================================================
Private Sub DrawActivePosition(ws As Worksheet, ticker As String, _
                                lots() As TILot, lotCount As Long)
    Dim totalShares As Double, totalCost As Double
    Dim i As Long
    For i = 1 To lotCount
        totalShares = totalShares + lots(i).Shares
        totalCost = totalCost + lots(i).Shares * lots(i).CostPerShare
    Next i

    Dim fx As Double: fx = GetFXToTWD(ticker)                   ' 31.6 for US, 1 for TW
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
        Call WriteValueMoney(ws, 4, 2, netExp, "NT$")
        Call WriteValuePrice(ws, 5, 2, wap)           ' per-share native (no fx)
    Else
        With ws.Cells(4, 2)
            .Value = "NO ACTIVE POSITION"
            .Font.Color = RGB(150, 150, 150)
            .Font.Italic = True
            .Interior.Color = RGB(0, 0, 0)
            .Borders.LineStyle = xlNone
        End With
        Call WriteDash(ws, 5, 2)
    End If

    If totalShares > 0 And lastPx > 0 Then
        Call WriteValuePrice(ws, 6, 2, lastPx)        ' per-share native (no fx)
        Call WriteValuePnL(ws, 7, 2, unrl, "NT$")
        Call WriteValuePct(ws, 8, 2, retPct)
    Else
        Call WriteDash(ws, 6, 2)
        Call WriteDash(ws, 7, 2)
        Call WriteDash(ws, 8, 2)
    End If
End Sub

' ================================================================
' SECTION 2: LIFETIME METRICS (columns D-E)
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

    Dim fx As Double: fx = GetFXToTWD(ticker)   ' 31.6 for US, 1 for TW

    ' --- Row 4: REALIZED PNL (TWD-converted) ---
    Call WriteValuePnL(ws, 4, 5, totalPnL * fx, "NT$")

    ' --- Row 5: LIFETIME EFF (Profit Factor) ---
    ' Defensive: clear borders/background, set NumberFormat BEFORE Value,
    ' and Round the value so even if format application fails the cell
    ' still displays 2 decimals max.
    With ws.Cells(5, 5)
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

    ' --- Row 6: VELOCITY/DAY (TWD-converted, PnL per holding day) ---
    Dim velocity As Double
    If totalDays > 0 Then velocity = totalPnL / totalDays
    Call WriteValuePnL(ws, 6, 5, velocity * fx, "NT$")
End Sub

' ================================================================
' SECTION 3: PROJECTION (columns G-H)
' H5 is the user-editable PRICE TARGET (yellow background)
' ================================================================
Private Sub DrawProjection(ws As Worksheet, ticker As String, _
                            lots() As TILot, lotCount As Long)
    Dim totalShares As Double
    Dim i As Long
    For i = 1 To lotCount: totalShares = totalShares + lots(i).Shares: Next i

    Dim lastPx As Double: lastPx = GetStockPriceSafe(ticker)
    Dim fx As Double: fx = GetFXToTWD(ticker)

    ' NOTE: labels drawn by DrawShellHeaders. Values only here.
    ' Row 4: LAST PRICE (per-share native)
    If lastPx > 0 Then
        Call WriteValuePrice(ws, 4, 8, lastPx)
    Else
        Call WriteDash(ws, 4, 8)
    End If

    ' Row 5: PRICE TARGET (per-share native input) - preserve existing value
    With ws.Cells(5, 8)
        If Not IsNumeric(.Value) Then .Value = vbNullString
    End With

    Dim target As Double
    If IsNumeric(ws.Cells(5, 8).Value) Then target = CDbl(ws.Cells(5, 8).Value)

    ' Row 6: PROJECTED PNL (TWD-converted)
    If totalShares > 0 And target > 0 And lastPx > 0 Then
        Dim projPnL As Double: projPnL = (target - lastPx) * totalShares * fx
        Call WriteValuePnL(ws, 6, 8, projPnL, "NT$")
    Else
        Call WriteDash(ws, 6, 8)
    End If

    ' Row 7: PROJECTED RETURN %
    If target > 0 And lastPx > 0 Then
        Dim projRet As Double: projRet = (target / lastPx) - 1
        Call WriteValuePct(ws, 7, 8, projRet)
    Else
        Call WriteDash(ws, 7, 8)
    End If
End Sub

' ================================================================
' SECTION 4: RECENT TRADE HISTORY (table, newest first)
' ================================================================
Private Sub DrawTradeHistory(ws As Worksheet, ticker As String, _
                              records() As TIRecord, recCount As Long)
    ' NOTE: Title (Row 10) and headers (Row 11) are drawn by DrawShellHeaders.
    ' We only fill data rows starting Row 12.
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

    Dim r As Long: r = 12
    For i = 1 To recCount
        With ws.Cells(r, 1)
            .NumberFormat = "yyyy-mm-dd"
            .Value = records(i).EntryDate
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlLeft
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 2)
            .NumberFormat = "yyyy-mm-dd"
            .Value = records(i).ExitDate
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlLeft
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 3)
            .Value = records(i).DaysHeld
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 4)
            .NumberFormat = "#,##0.##"
            .Value = records(i).Shares
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlRight
            .Borders.LineStyle = xlNone
        End With
        ' --- PnL converted to TWD ---
        Dim pnlTWD As Double: pnlTWD = records(i).PnL * fx
        With ws.Cells(r, 5)
            .NumberFormat = pnlFmtStr
            .Value = pnlTWD
            .Font.Color = USPnLColor(pnlTWD)
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 6)
            .NumberFormat = "0.00%"
            .Value = Round(records(i).ReturnPct, 4)
            .Font.Color = USPnLColor(records(i).PnL)
            .HorizontalAlignment = xlRight
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 7)
            .Value = records(i).Broker
            .Font.Color = RGB(150, 150, 150)
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
            .Borders.LineStyle = xlNone
        End With
        ws.Rows(r).RowHeight = 16
        r = r + 1
    Next i
End Sub

' ================================================================
' CELL HELPERS
' ================================================================
Private Sub WriteLabel(ws As Worksheet, r As Long, c As Long, txt As String)
    With ws.Cells(r, c)
        .Value = txt
        .Font.Color = RGB(0, 200, 255)
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
    End With
End Sub

' All value-writers below set NumberFormat BEFORE Value and explicitly
' clear borders/background to prevent stale formatting from prior renders.
'
' Currency symbol "NT$" must be quoted in the format string. Bare
' "$" in Excel format codes is treated as a locale-aware currency
' specifier (in zh-TW it expands to "NT$"), so an unquoted "NT$..."
' was being mis-parsed and silently rejected on some Excel versions.
' Using Chr(34) (double quote) around the symbol forces literal text.
Private Sub WriteValueMoney(ws As Worksheet, r As Long, c As Long, _
                             v As Double, sym As String)
    With ws.Cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = MoneyFmt(sym)
        .Value = Round(v, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub WriteValuePrice(ws As Worksheet, r As Long, c As Long, v As Double)
    With ws.Cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = "#,##0.00"
        .Value = Round(v, 2)
        .Font.Color = RGB(221, 221, 221)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub WriteValuePnL(ws As Worksheet, r As Long, c As Long, _
                           v As Double, sym As String)
    With ws.Cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = PnLFmt(sym)
        .Value = Round(v, 0)
        .Font.Color = USPnLColor(v)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
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
    With ws.Cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = "0.00%"
        .Value = Round(v, 4)
        .Font.Color = USPnLColor(v)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub WriteDash(ws As Worksheet, r As Long, c As Long)
    With ws.Cells(r, c)
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

' Suffix-based classification:
'   .TW or .TWO suffix -> TW (no FX conversion)
'   Everything else    -> US (FX x 31.6 applied to aggregates)
'
' This is a STRICT rule. Special codes like "00981A" without a suffix
' are treated as US. To classify a TW security correctly, the ticker
' MUST be typed with its full suffix (e.g. "00981A.TWO", "2330.TW").
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
'   US ticker -> FX_USD_TWD (31.6)
Private Function GetFXToTWD(ticker As String) As Double
    If IsTWTicker(ticker) Then GetFXToTWD = 1 Else GetFXToTWD = FX_USD_TWD
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

