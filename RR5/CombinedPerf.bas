Attribute VB_Name = "CombinedPerf"
Option Explicit

' =============================================================
' COMBINED PERFORMANCE TABLE  v1.1
' RR4 (Stocks) + RR5 (Derivatives) joint summary
'
' Entry points:
'   BuildCombinedPerf   -- full rebuild (format + data)
'   RefreshCombinedPerf -- data refresh only (fast)
'   SyncRR4Data         -- re-read open RR4 workbook
'
' RR5 data: auto-read from this workbook (B1/F2/H2/P1/L1/J2/L2/N2)
' RR4 data: auto-search open workbook with "RR4" in name;
'           orange cells if not found -- type manually then Refresh
' =============================================================

Private Const SH_OUT  As String = "CombinedPerf"
Private Const SH_RR5  As String = "RR5"
Private Const SH_LOG  As String = "CPLog"

' Column in RR4 Analysis sheet that holds per-ticker market value.
' Run DiagnoseRR4 to inspect the sheet layout, then adjust if needed.
Private Const RR4_PIE_VAL_COL As Long = 7   ' col G = NET EXPOS in RR4 Analysis

' Color palette (dark Bloomberg style)
Private Const LOG_START As Date = #1/1/2026#     ' CombinedLog date floor

Private Const CLR_BG_DARK   As Long = 1973790    ' #1E1E1E = RGB(30,30,30) -- CPLog sheet bg
Private Const CLR_BG_MID    As Long = 1776411    ' #1B1B1B = RGB(27,27,27) -- dark alternating row
Private Const CLR_BG_LIGHT  As Long = 1973790    ' #1E1E1E = RGB(30,30,30)
Private Const CLR_SECTION   As Long = 1973790    ' #1E1E1E = RGB(30,30,30) -- gray alternating row
Private Const CLR_FG_WHITE  As Long = 16777215
Private Const CLR_FG_SILVER As Long = 12895428   ' #C4C4C4
Private Const CLR_FG_GOLD   As Long = 55295      ' #FFD700 = RGB(255,215,0) -- actual gold
Private Const CLR_FG_CYAN   As Long = 6750105    ' #66CCFF approx
Private Const CLR_FG_GREEN  As Long = 6479952    ' #62DC50 approx (RR5 header color)
Private Const CLR_BORDER    As Long = 5066061    ' #4D4D4D
Private Const CLR_ORANGE_BG As Long = 4687923    ' manual-input hint

' CPLog group colors: RR4 = yellow/gold, RR5 = blue, Combined = gray.
' Applied to header text, line-chart series, and line-chart titles alike.
Private Const CLR_GRP_RR4  As Long = 55295      ' RGB(255,215,0) -- reuses tested gold/yellow
Private Const CLR_GRP_RR5  As Long = 16754010   ' RGB(90,165,255) -- blue
Private Const CLR_GRP_COMB As Long = 12895428   ' RGB(196,196,196) -- gray (same as CLR_FG_SILVER)

' RR5 sheet cells for IV Avg and TWII Price -- adjust if layout differs
Private Const RR5_IV_CELL   As String = "R2"    ' avg implied vol of options portfolio
Private Const RR5_TWII_CELL As String = "T2"    ' TWII index price

' Row constants
Private Const R_TITLE    As Long = 1
Private Const R_UPDATED  As Long = 2
Private Const R_HDR      As Long = 4
Private Const R_NEXPOS   As Long = 5
Private Const R_UNRL_PNL As Long = 6
Private Const R_RL_PNL   As Long = 7
Private Const R_COMB_RL  As Long = 8   ' Combined RL PNL -- adjacent to Realized PNL
Private Const R_TOT_PNL  As Long = 9
Private Const R_UNRL_PCT As Long = 10
Private Const R_POSCOUNT As Long = 11
Private Const R_SEP1     As Long = 12
Private Const R_D_HDR    As Long = 13
Private Const R_DELTA    As Long = 14
Private Const R_THETA    As Long = 15
Private Const R_VEGA     As Long = 16
Private Const R_MARGIN   As Long = 17
Private Const R_SEP2     As Long = 18
Private Const R_BK_HDR   As Long = 19
Private Const R_BK_COL   As Long = 20
Private Const R_FUT      As Long = 21
Private Const R_OPT      As Long = 22
Private Const R_WRT      As Long = 23
Private Const R_SEP3     As Long = 24
Private Const R_NOTE     As Long = 25

' Col constants
Private Const C_LABEL As Long = 1   ' A
Private Const C_RR4   As Long = 2   ' B
Private Const C_RR5   As Long = 3   ' C
Private Const C_TOTAL As Long = 4   ' D


' =============================================================
' PUBLIC ENTRY POINTS
' =============================================================

Public Sub BuildCombinedPerf()
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = EnsureSheet()
    Call DrawShell(ws)
    Call RefreshValues(ws)

    Application.ScreenUpdating = True
    Application.StatusBar = "CombinedPerf built OK  " & Format(Now, "hh:mm:ss")
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "BuildCombinedPerf failed: " & Err.Description, vbExclamation, "CombinedPerf"
End Sub

Public Sub RefreshCombinedPerf()
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_OUT)
    On Error GoTo Fail
    If ws Is Nothing Then Call BuildCombinedPerf: Exit Sub

    Call RefreshValues(ws)

    Application.ScreenUpdating = True
    Application.StatusBar = "CombinedPerf refreshed  " & Format(Now, "hh:mm:ss")
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "RefreshCombinedPerf failed: " & Err.Description, vbExclamation, "CombinedPerf"
End Sub

Public Sub SyncRR4Data()
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_OUT)
    On Error GoTo Fail
    If ws Is Nothing Then Call BuildCombinedPerf: Exit Sub

    Dim found As Boolean
    Dim rr4 As Variant
    rr4 = GetRR4Vals(found)
    Call WriteRR4Cells(ws, rr4, found)
    Call RecalcTotals(ws)
    ws.Cells(R_UPDATED, C_LABEL).Value = "Updated: " & Format(Now, "yyyy/mm/dd hh:mm:ss")
    Call AppendLog(ws)

    Application.ScreenUpdating = True
    If found Then
        Application.StatusBar = "RR4 synced OK  " & Format(Now, "hh:mm:ss")
    Else
        MsgBox "RR4 workbook not found.  Open it first, or type values in orange cells.", _
               vbInformation, "RR4 not found"
    End If
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "SyncRR4Data failed: " & Err.Description, vbExclamation, "CombinedPerf"
End Sub

' Single consolidated "Refresh All" entry point -- the sheet's only button.
' Runs the full daily routine in one click: data refresh (self-heals into
' BuildCombinedPerf if the sheet doesn't exist yet), both holdings pie
' charts, then the CPLog history/chart/risk panel (last, so it picks up
' the numbers this refresh just wrote). One-off tools that don't belong in
' a routine refresh -- SyncRR4Data, ImportRR4History, a from-scratch
' BuildCombinedPerf -- stay callable via Alt+F8, just not on a button.
Public Sub RefreshAllCombined()
    Call RefreshCombinedPerf
    Call BuildHoldingsPieCharts
    Call BuildCombinedHistoryLog
    Application.StatusBar = "CombinedPerf: full refresh done  " & Format(Now, "hh:mm:ss")
End Sub


' =============================================================
' PRIVATE: ensure sheet
' =============================================================
Private Function EnsureSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_OUT)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SH_OUT
    Else
        ws.Cells.Clear
        ws.Cells.ClearFormats
        On Error Resume Next
        ws.DrawingObjects.Delete
        On Error GoTo 0
    End If
    Set EnsureSheet = ws
End Function


' =============================================================
' PRIVATE: draw static shell (labels, formatting, buttons)
' =============================================================
Private Sub DrawShell(ws As Worksheet)
    ws.Cells.Interior.Color = CLR_BG_LIGHT
    ws.Cells.Font.Color = CLR_FG_SILVER
    ' DisplayGridlines belongs to Window, not Worksheet -- set via ActiveWindow after ws.Activate

    ' column widths
    ws.Columns(C_LABEL).ColumnWidth = 26
    ws.Columns(C_RR4).ColumnWidth   = 18
    ws.Columns(C_RR5).ColumnWidth   = 18
    ws.Columns(C_TOTAL).ColumnWidth = 18
    ws.Columns(5).ColumnWidth = 14

    ' -- Row 1: title --
    ws.Rows(R_TITLE).RowHeight = 28
    With ws.Range(ws.Cells(R_TITLE, C_LABEL), ws.Cells(R_TITLE, C_TOTAL))
        .Merge
        .Value = "RR4 + RR5  Joint Performance"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = CLR_FG_WHITE
        .Interior.Color = CLR_BG_DARK
        .HorizontalAlignment = xlCenter
    End With

    ' -- Row 2: updated timestamp --
    ws.Rows(R_UPDATED).RowHeight = 15
    With ws.Range(ws.Cells(R_UPDATED, C_LABEL), ws.Cells(R_UPDATED, C_TOTAL))
        .Merge
        .Font.Size = 9
        .Font.Color = RGB(130, 130, 130)
        .Interior.Color = CLR_BG_DARK
        .HorizontalAlignment = xlRight
    End With

    ' -- Row 3: spacer --
    ws.Rows(3).RowHeight = 5

    ' -- Row 4: column headers --
    ws.Rows(R_HDR).RowHeight = 20
    Dim hLabels As Variant
    hLabels = Array("Metric", "RR4  Stocks", "RR5  Derivatives", "Combined")
    Dim hColors As Variant
    hColors = Array(CLR_FG_WHITE, CLR_FG_GOLD, CLR_FG_CYAN, RGB(180, 255, 130))
    Dim c As Long
    For c = C_LABEL To C_TOTAL
        With ws.Cells(R_HDR, c)
            .Value = hLabels(c - 1)
            .Font.Bold = True
            .Font.Color = hColors(c - 1)
            .Interior.Color = CLR_SECTION
            .HorizontalAlignment = xlCenter
        End With
    Next c

    ' -- Rows 5-7: Net Exposure / Unrealized PNL / Realized PNL --
    Dim mLabels1 As Variant
    mLabels1 = Array("Net Exposure (TWD)", "Unrealized PNL (TWD)", "Realized PNL (TWD)")
    Dim r As Long, i As Long
    For i = 0 To 2
        r = R_NEXPOS + i
        ws.Rows(r).RowHeight = 19
        With ws.Cells(r, C_LABEL)
            .Value = mLabels1(i)
            .Interior.Color = CLR_BG_MID
        End With
        ws.Cells(r, C_RR4).Interior.Color = CLR_BG_LIGHT
        ws.Cells(r, C_RR5).Interior.Color = CLR_BG_LIGHT
        With ws.Cells(r, C_TOTAL)
            .Interior.Color = CLR_BG_MID
            .Font.Bold = True
            .Font.Color = RGB(180, 255, 130)
        End With
    Next i

    ' -- Row 8: Combined RL PNL (adjacent to Realized PNL, same type) --
    ws.Rows(R_COMB_RL).RowHeight = 17
    With ws.Cells(R_COMB_RL, C_LABEL)
        .Value = "  Combined RL PNL"
        .Interior.Color = CLR_BG_MID
    End With
    ws.Cells(R_COMB_RL, C_RR4).Interior.Color = CLR_BG_LIGHT
    ws.Cells(R_COMB_RL, C_RR5).Interior.Color = CLR_BG_LIGHT
    With ws.Cells(R_COMB_RL, C_TOTAL)
        .Interior.Color = CLR_BG_MID
        .Font.Bold = True
        .Font.Color = RGB(180, 255, 130)
    End With

    ' -- Rows 9-11: Total PNL / Realized Return % / Positions --
    Dim mLabels2 As Variant
    mLabels2 = Array("Total PNL (TWD)", "Realized Return %", "Positions")
    Dim mRows2 As Variant
    mRows2 = Array(R_TOT_PNL, R_UNRL_PCT, R_POSCOUNT)
    For i = 0 To 2
        r = mRows2(i)
        ws.Rows(r).RowHeight = 19
        With ws.Cells(r, C_LABEL)
            .Value = mLabels2(i)
            .Interior.Color = CLR_BG_MID
        End With
        ws.Cells(r, C_RR4).Interior.Color = CLR_BG_LIGHT
        ws.Cells(r, C_RR5).Interior.Color = CLR_BG_LIGHT
        With ws.Cells(r, C_TOTAL)
            .Interior.Color = CLR_BG_MID
            .Font.Bold = True
            .Font.Color = RGB(180, 255, 130)
        End With
    Next i

    Call PutSep(ws, R_SEP1)
    Call PutSecHdr(ws, R_D_HDR, "-- RR5 Derivatives Detail --")

    ' -- Rows 13-16: derivatives-only metrics --
    Dim dLabels As Variant
    dLabels = Array("Net Delta", "Net Theta (TWD/day)", "Net Vega (TWD)", "Margin Requirement (TWD)")
    For i = 0 To 3
        r = R_DELTA + i
        ws.Rows(r).RowHeight = 19
        With ws.Cells(r, C_LABEL)
            .Value = dLabels(i)
            .Interior.Color = CLR_BG_MID
        End With
        ws.Cells(r, C_RR4).Interior.Color = RGB(25, 25, 25)  ' n/a for stocks
        ws.Cells(r, C_RR5).Interior.Color = CLR_BG_LIGHT
        ws.Cells(r, C_TOTAL).Interior.Color = CLR_BG_LIGHT
    Next i

    Call PutSep(ws, R_SEP2)
    Call PutSecHdr(ws, R_BK_HDR, "-- RR5 Breakdown by Type --")

    ' -- Row 19: breakdown headers --
    ws.Rows(R_BK_COL).RowHeight = 17
    Dim bkH As Variant
    bkH = Array("Type", "Net Exposure (TWD)", "Unrealized PNL (TWD)", "Net Delta")
    For c = C_LABEL To C_TOTAL
        With ws.Cells(R_BK_COL, c)
            .Value = bkH(c - 1)
            .Font.Bold = True
            .Font.Color = RGB(170, 170, 170)
            .Interior.Color = RGB(38, 38, 38)
            .HorizontalAlignment = xlCenter
        End With
    Next c

    ' -- Rows 20-22: breakdown rows --
    Dim bkLabels As Variant
    bkLabels = Array("Futures", "Options", "Warrants")
    For i = 0 To 2
        r = R_FUT + i
        ws.Rows(r).RowHeight = 19
        ws.Cells(r, C_LABEL).Value = bkLabels(i)
        ws.Cells(r, C_LABEL).Interior.Color = CLR_BG_MID
        ws.Cells(r, C_RR4).Interior.Color = CLR_BG_LIGHT
        ws.Cells(r, C_RR5).Interior.Color = CLR_BG_LIGHT
        ws.Cells(r, C_TOTAL).Interior.Color = CLR_BG_LIGHT
    Next i

    Call PutSep(ws, R_SEP3)

    ' -- Row 24: note --
    ws.Rows(R_NOTE).RowHeight = 28
    With ws.Range(ws.Cells(R_NOTE, C_LABEL), ws.Cells(R_NOTE, C_TOTAL))
        .Merge
        .Value = "Note: Orange cells = manual input.  Or open RR4 workbook first, then click [Refresh All].  " & _
                 "Realized PNL reads from Realized sheet column I."
        .Font.Size = 8
        .Font.Color = RGB(130, 130, 130)
        .WrapText = True
    End With

    ' -- Button (column E area) --
    Dim btnTop As Double: btnTop = ws.Rows(R_TITLE).Top + 2
    Dim btnH As Double: btnH = ws.Rows(R_TITLE).Height + ws.Rows(R_UPDATED).Height - 4
    Dim btnL As Double: btnL = ws.Cells(R_TITLE, 5).Left + 4

    Dim btn As Object
    Set btn = ws.Buttons.Add(btnL, btnTop, 110, btnH)
    btn.Caption = "Refresh All": btn.OnAction = "RefreshAllCombined"

    ' -- Number formats --
    Call ApplyFormats(ws)

    ' -- Borders --
    Call PutBorder(ws, R_HDR, R_POSCOUNT)
    Call PutBorder(ws, R_D_HDR, R_MARGIN)
    Call PutBorder(ws, R_BK_COL, R_WRT)
End Sub


' =============================================================
' PRIVATE: refresh all values
' =============================================================
Private Sub RefreshValues(ws As Worksheet)
    ' Force a full recalc BEFORE reading RR5. RefreshValues can now run from
    ' RR5_Core.RebuildPositionsFromTransactions while a transaction is still
    ' being saved -- at that point Application.Calculation is pinned to
    ' xlCalculationManual by ShowTransactionForm for the whole modal session,
    ' so RR5's summary formulas (B1/F2/H2/N15/Q27/Q39...) may not yet reflect
    ' the position data that was just migrated in. Without this, CombinedPerf/
    ' CPLog capture a stale (pre-transaction) snapshot even though RR5 itself
    ' reads correctly once Excel later recalculates automatically.
    Application.Calculate

    Dim rr5 As Variant: rr5 = GetRR5Vals()
    Call WriteRR5Cells(ws, rr5)

    Dim found As Boolean
    Dim rr4 As Variant: rr4 = GetRR4Vals(found)
    Call WriteRR4Cells(ws, rr4, found)

    Call RecalcTotals(ws)
    ws.Cells(R_UPDATED, C_LABEL).Value = "Updated: " & Format(Now, "yyyy/mm/dd hh:mm:ss")

    Call AppendLog(ws)
End Sub


' =============================================================
' PRIVATE: read RR5 values from this workbook
' =============================================================
Private Function GetRR5Vals() As Variant
    ' idx: 0=NetExpos 1=UnrlPNL 2=RlPNL 3=UnrlPct 4=Positions
    '      5=Delta 6=Theta 7=Vega 8=Margin
    '      9=FutExpos 10=FutPNL 11=FutDelta
    '      12=OptExpos 13=OptPNL 14=OptDelta
    '      15=WrtExpos 16=WrtPNL 17=WrtDelta
    Dim v(0 To 17) As Variant
    Dim i As Long
    For i = 0 To 17: v(i) = 0: Next i

    Dim wsR As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets(SH_RR5)
    On Error GoTo 0
    If wsR Is Nothing Then GetRR5Vals = v: Exit Function

    On Error Resume Next
    v(0) = V2D(wsR.Range("B1").Value)   ' Total MV
    v(1) = V2D(wsR.Range("F2").Value)   ' Unrealized PNL
    v(2) = V2D(wsR.Range("H2").Value)   ' Realized PNL (=SUM Realized!I)
    v(3) = V2D(wsR.Range("D2").Value)   ' Unrealized PNL %
    v(4) = V2D(wsR.Range("P1").Value)   ' Positions
    v(5) = V2D(wsR.Range("L1").Value)   ' Net Delta
    v(6) = V2D(wsR.Range("J2").Value)   ' Net Theta
    v(7) = V2D(wsR.Range("L2").Value)   ' Net Vega
    v(8) = V2D(wsR.Range("N2").Value)   ' Margin

    v(9)  = V2D(wsR.Range("N15").Value)  ' Futures Net Expos
    v(10) = V2D(wsR.Range("O15").Value)  ' Futures Unrealized PNL
    v(11) = V2D(wsR.Range("S15").Value)  ' Futures Net Delta

    v(12) = V2D(wsR.Range("Q27").Value)  ' Options Net Expos
    v(13) = V2D(wsR.Range("R27").Value)  ' Options Unrealized PNL
    v(14) = V2D(wsR.Range("Y27").Value)  ' Options Net Delta

    v(15) = V2D(wsR.Range("Q39").Value)  ' Warrants Net Expos
    v(16) = V2D(wsR.Range("R39").Value)  ' Warrants Unrealized PNL
    v(17) = V2D(wsR.Range("X39").Value)  ' Warrants Net Delta (SUMPRODUCT)
    On Error GoTo 0

    GetRR5Vals = v
End Function


' =============================================================
' PRIVATE: read RR4 values from open workbook
' =============================================================
Private Function GetRR4Vals(ByRef found As Boolean) As Variant
    ' idx: 0=NetExpos 1=UnrlPNL 2=RlPNL 3=UnrlPct 4=Positions
    Dim v(0 To 4) As Variant
    Dim i As Long
    For i = 0 To 4: v(i) = 0: Next i
    found = False

    Dim wb As Workbook
    Set wb = FindRR4WB()
    If wb Is Nothing Then GetRR4Vals = v: Exit Function
    found = True

    On Error Resume Next
    ' Read from RR4 Analysis sheet -- fixed cell layout confirmed by DiagnoseRR4:
    '   R5  col B = Total Return %
    '   R6  col B = Total Cost (cost basis)
    '   R7  col B = Unrealized PnL
    '   R8  col B = Realized PnL
    '   R9  col B = Total PnL
    '   R31+ col A = Ticker (individual positions; col C = shares numeric)
    Dim wsA As Worksheet
    Set wsA = wb.Sheets("Analysis")
    If Not wsA Is Nothing Then
        Dim totalCost As Double: totalCost = V2D(wsA.Cells(6, 2).Value)
        v(1) = V2D(wsA.Cells(7, 2).Value)                         ' Unrealized PnL
        v(2) = V2D(wsA.Cells(8, 2).Value)                         ' Realized PnL
        v(0) = totalCost + v(1)                                    ' Net Exposure = Cost + Unrealized
        v(3) = v(2) / 550000                                       ' Realized Return % = Realized PNL / 550k

        ' Count positions: find "Ticker" header in col A, count non-empty rows below
        Dim r As Long, tk As String
        Dim tickerHdr As Long: tickerHdr = 0
        For r = 1 To 60
            If InStr(1, LCase(Trim(CStr(wsA.Cells(r, 1).Value))), "ticker") > 0 Then
                tickerHdr = r: Exit For
            End If
        Next r
        If tickerHdr = 0 Then tickerHdr = 26   ' fallback: tickers start below row 26
        v(4) = 0
        For r = tickerHdr + 1 To 300
            tk = Trim(CStr(wsA.Cells(r, 1).Value))
            If tk <> "" Then v(4) = v(4) + 1
        Next r
    End If
    On Error GoTo 0

    GetRR4Vals = v
End Function


Private Function FindRR4WB() As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            If InStr(UCase(wb.Name), "RR4") > 0 Then
                Set FindRR4WB = wb
                Exit Function
            End If
        End If
    Next wb
    Set FindRR4WB = Nothing
End Function


' =============================================================
' PRIVATE: write RR5 cells
' =============================================================
Private Sub WriteRR5Cells(ws As Worksheet, v As Variant)
    ws.Cells(R_NEXPOS,   C_RR5).Value = v(0)
    ws.Cells(R_UNRL_PNL, C_RR5).Value = v(1)
    ws.Cells(R_RL_PNL,   C_RR5).Value = v(2)
    ws.Cells(R_TOT_PNL,  C_RR5).Value = v(1) + v(2)
    ws.Cells(R_UNRL_PCT, C_RR5).Value = v(3)
    ws.Cells(R_POSCOUNT, C_RR5).Value = v(4)

    ws.Cells(R_DELTA,  C_RR5).Value = v(5)
    ws.Cells(R_THETA,  C_RR5).Value = v(6)
    ws.Cells(R_VEGA,   C_RR5).Value = v(7)
    ws.Cells(R_MARGIN, C_RR5).Value = v(8)

    ws.Cells(R_DELTA,  C_TOTAL).Value = v(5)
    ws.Cells(R_THETA,  C_TOTAL).Value = v(6)
    ws.Cells(R_VEGA,   C_TOTAL).Value = v(7)
    ws.Cells(R_MARGIN, C_TOTAL).Value = v(8)

    ' Breakdown (use C_RR4 col as Exposure, C_RR5 as PNL, C_TOTAL as Delta)
    ws.Cells(R_FUT, C_RR4).Value = v(9):  ws.Cells(R_FUT, C_RR5).Value = v(10): ws.Cells(R_FUT, C_TOTAL).Value = v(11)
    ws.Cells(R_OPT, C_RR4).Value = v(12): ws.Cells(R_OPT, C_RR5).Value = v(13): ws.Cells(R_OPT, C_TOTAL).Value = v(14)
    ws.Cells(R_WRT, C_RR4).Value = v(15): ws.Cells(R_WRT, C_RR5).Value = v(16): ws.Cells(R_WRT, C_TOTAL).Value = v(17)

    ColorPnl ws.Cells(R_UNRL_PNL, C_RR5), v(1)
    ColorPnl ws.Cells(R_RL_PNL,   C_RR5), v(2)
    ColorPnl ws.Cells(R_TOT_PNL,  C_RR5), v(1) + v(2)
    ColorPnl ws.Cells(R_UNRL_PCT, C_RR5), v(3)
    ColorPnl ws.Cells(R_THETA,    C_RR5), v(6)
End Sub


' =============================================================
' PRIVATE: write RR4 cells
' =============================================================
Private Sub WriteRR4Cells(ws As Worksheet, v As Variant, found As Boolean)
    Dim bgClr As Long
    If found Then bgClr = CLR_BG_LIGHT Else bgClr = CLR_ORANGE_BG

    Dim rows4 As Variant
    rows4 = Array(R_NEXPOS, R_UNRL_PNL, R_RL_PNL, R_TOT_PNL, R_UNRL_PCT, R_POSCOUNT)
    Dim i As Long
    For i = 0 To 5
        With ws.Cells(rows4(i), C_RR4)
            .Interior.Color = bgClr
            If found Then
                Select Case i
                    Case 0: .Value = v(0)
                    Case 1: .Value = v(1)
                    Case 2: .Value = v(2)
                    Case 3: .Value = v(1) + v(2)
                    Case 4: .Value = v(3)
                    Case 5: .Value = v(4)
                End Select
            Else
                If .Value = 0 Or .Value = "" Then .Value = ""
                .Font.Color = RGB(255, 180, 80)
            End If
        End With
    Next i

    If found Then
        ColorPnl ws.Cells(R_UNRL_PNL, C_RR4), v(1)
        ColorPnl ws.Cells(R_RL_PNL,   C_RR4), v(2)
        ColorPnl ws.Cells(R_TOT_PNL,  C_RR4), v(1) + v(2)
        ColorPnl ws.Cells(R_UNRL_PCT, C_RR4), v(3)
    End If
End Sub


' =============================================================
' PRIVATE: recalculate totals column
' =============================================================
Private Sub RecalcTotals(ws As Worksheet)
    Dim a As Double, b As Double

    ' Net Exposure
    a = SN(ws.Cells(R_NEXPOS, C_RR4)): b = SN(ws.Cells(R_NEXPOS, C_RR5))
    ws.Cells(R_NEXPOS, C_TOTAL).Value = a + b

    ' Unrealized PNL
    a = SN(ws.Cells(R_UNRL_PNL, C_RR4)): b = SN(ws.Cells(R_UNRL_PNL, C_RR5))
    ws.Cells(R_UNRL_PNL, C_TOTAL).Value = a + b
    ColorPnl ws.Cells(R_UNRL_PNL, C_TOTAL), a + b

    ' Realized PNL
    a = SN(ws.Cells(R_RL_PNL, C_RR4)): b = SN(ws.Cells(R_RL_PNL, C_RR5))
    ws.Cells(R_RL_PNL, C_TOTAL).Value = a + b
    ColorPnl ws.Cells(R_RL_PNL, C_TOTAL), a + b

    ' Total PNL
    a = SN(ws.Cells(R_TOT_PNL, C_RR4)): b = SN(ws.Cells(R_TOT_PNL, C_RR5))
    ws.Cells(R_TOT_PNL, C_TOTAL).Value = a + b
    ColorPnl ws.Cells(R_TOT_PNL, C_TOTAL), a + b

    ' Realized Return % (combined) = total Realized PNL / 550,000
    Dim rl4 As Double: rl4 = SN(ws.Cells(R_RL_PNL, C_RR4))
    Dim rl5 As Double: rl5 = SN(ws.Cells(R_RL_PNL, C_RR5))
    Dim combRet As Double: combRet = (rl4 + rl5) / 550000
    ws.Cells(R_UNRL_PCT, C_TOTAL).Value = combRet
    ColorPnl ws.Cells(R_UNRL_PCT, C_TOTAL), combRet

    ' Positions
    ws.Cells(R_POSCOUNT, C_TOTAL).Value = SN(ws.Cells(R_POSCOUNT, C_RR4)) + _
                                           SN(ws.Cells(R_POSCOUNT, C_RR5))

    ' Combined RL PNL (adjacent to Realized PNL)
    Dim combRL As Double
    combRL = SN(ws.Cells(R_RL_PNL, C_RR4)) + SN(ws.Cells(R_RL_PNL, C_RR5))
    ws.Cells(R_COMB_RL, C_TOTAL).Value = combRL
    ColorPnl ws.Cells(R_COMB_RL, C_TOTAL), combRL
End Sub


' =============================================================
' PRIVATE: number formats
' =============================================================
Private Sub ApplyFormats(ws As Worksheet)
    Dim twdFmt As String: twdFmt = "#,##0;[Red]-#,##0"
    Dim pctFmt As String: pctFmt = "0.00%;[Red]-0.00%"
    Dim dltFmt As String: dltFmt = "#,##0.0000;[Red]-#,##0.0000"
    Dim intFmt As String: intFmt = "0"

    Dim r As Long
    For r = R_NEXPOS To R_RL_PNL
        ws.Cells(r, C_RR4).NumberFormat = twdFmt
        ws.Cells(r, C_RR5).NumberFormat = twdFmt
        ws.Cells(r, C_TOTAL).NumberFormat = twdFmt
    Next r
    ws.Cells(R_COMB_RL, C_TOTAL).NumberFormat = twdFmt
    ws.Cells(R_TOT_PNL, C_RR4).NumberFormat = twdFmt
    ws.Cells(R_TOT_PNL, C_RR5).NumberFormat = twdFmt
    ws.Cells(R_TOT_PNL, C_TOTAL).NumberFormat = twdFmt
    ws.Cells(R_UNRL_PCT, C_RR4).NumberFormat = pctFmt
    ws.Cells(R_UNRL_PCT, C_RR5).NumberFormat = pctFmt
    ws.Cells(R_UNRL_PCT, C_TOTAL).NumberFormat = pctFmt
    ws.Cells(R_POSCOUNT, C_RR4).NumberFormat = intFmt
    ws.Cells(R_POSCOUNT, C_RR5).NumberFormat = intFmt
    ws.Cells(R_POSCOUNT, C_TOTAL).NumberFormat = intFmt

    ws.Cells(R_DELTA, C_RR5).NumberFormat = dltFmt
    ws.Cells(R_DELTA, C_TOTAL).NumberFormat = dltFmt
    ws.Cells(R_THETA, C_RR5).NumberFormat = twdFmt
    ws.Cells(R_THETA, C_TOTAL).NumberFormat = twdFmt
    ws.Cells(R_VEGA, C_RR5).NumberFormat = twdFmt
    ws.Cells(R_VEGA, C_TOTAL).NumberFormat = twdFmt
    ws.Cells(R_MARGIN, C_RR5).NumberFormat = twdFmt
    ws.Cells(R_MARGIN, C_TOTAL).NumberFormat = twdFmt

    Dim bkRows As Variant: bkRows = Array(R_FUT, R_OPT, R_WRT)
    Dim i As Long
    For i = 0 To 2
        ws.Cells(bkRows(i), C_RR4).NumberFormat = twdFmt
        ws.Cells(bkRows(i), C_RR5).NumberFormat = twdFmt
        ws.Cells(bkRows(i), C_TOTAL).NumberFormat = dltFmt
    Next i
End Sub


' =============================================================
' PRIVATE: helpers
' =============================================================

Private Sub PutSep(ws As Worksheet, r As Long)
    ws.Rows(r).RowHeight = 7
    With ws.Range(ws.Cells(r, C_LABEL), ws.Cells(r, C_TOTAL))
        .Interior.Color = CLR_BG_DARK
    End With
End Sub

Private Sub PutSecHdr(ws As Worksheet, r As Long, txt As String)
    ws.Rows(r).RowHeight = 17
    With ws.Range(ws.Cells(r, C_LABEL), ws.Cells(r, C_TOTAL))
        .Merge
        .Value = txt
        .Interior.Color = CLR_SECTION
        .Font.Color = RGB(160, 210, 255)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub PutBorder(ws As Worksheet, rStart As Long, rEnd As Long)
    With ws.Range(ws.Cells(rStart, C_LABEL), ws.Cells(rEnd, C_TOTAL))
        Dim sides As Variant
        sides = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                      xlInsideVertical, xlInsideHorizontal)
        Dim s As Long
        For s = 0 To UBound(sides)
            On Error Resume Next
            .Borders(sides(s)).LineStyle = xlContinuous
            .Borders(sides(s)).Color = CLR_BORDER
            On Error GoTo 0
        Next s
    End With
End Sub

Private Sub ColorPnl(c As Range, ByVal v As Double)
    If v > 0 Then
        c.Font.Color = RGB(80, 220, 100)
    ElseIf v < 0 Then
        c.Font.Color = RGB(255, 80, 80)
    Else
        c.Font.Color = CLR_FG_SILVER
    End If
End Sub

Private Function V2D(v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then V2D = CDbl(v) Else V2D = 0
    On Error GoTo 0
End Function

Private Function SN(c As Range) As Double
    On Error Resume Next
    If IsNumeric(c.Value) Then SN = CDbl(c.Value) Else SN = 0
    On Error GoTo 0
End Function


' =============================================================
' PUBLIC: open history log sheet + rebuild line chart
' =============================================================
Public Sub BuildCombinedHistoryLog()
    Dim wl As Worksheet: Set wl = EnsureLogSheet()
    ' Ensure sheet base background is dark gray (not pure black)
    wl.Cells.Interior.Color = CLR_BG_DARK

    ' --- Remove rows before LOG_START ---
    Dim lr As Long: lr = wl.Cells(wl.Rows.Count, 1).End(xlUp).Row
    Application.ScreenUpdating = False
    Dim dr As Long
    Dim cellDt As Date
    For dr = lr To 3 Step -1
        If IsDate(wl.Cells(dr, 1).Value) Then
            cellDt = CDate(wl.Cells(dr, 1).Value)
            If cellDt < LOG_START Then wl.Rows(dr).Delete
        End If
    Next dr
    lr = wl.Cells(wl.Rows.Count, 1).End(xlUp).Row

    ' --- Deduplicate: for same-date rows keep only the LAST (most complete) row ---
    For dr = lr To 4 Step -1   ' scan bottom-up so deleting doesn't shift index
        If wl.Cells(dr, 1).Value = "" Then GoTo NextDR
        Dim dupR As Long: dupR = 0
        Dim sr As Long
        For sr = dr - 1 To 3 Step -1
            If Int(wl.Cells(sr, 1).Value) = Int(wl.Cells(dr, 1).Value) Then
                dupR = sr: Exit For
            End If
        Next sr
        If dupR > 0 Then wl.Rows(dupR).Delete  ' delete the earlier duplicate
NextDR:
    Next dr
    Application.ScreenUpdating = True
    lr = wl.Cells(wl.Rows.Count, 1).End(xlUp).Row   ' refresh after dedup

    ' Re-stamp alternating gray/black on every data row (CLR_SECTION / CLR_BG_MID)
    Dim rr As Long
    Dim bRL As Double, bRet As Double, bComb As Double
    For rr = 3 To lr
        wl.Rows(rr).Interior.Color = IIf(rr Mod 2 = 1, CLR_SECTION, CLR_BG_MID)
        ' Backfill col 5 (RR4 RL Ret%) -- derive from col 4 / 550k if empty
        bRL = V2D(wl.Cells(rr, 4).Value)
        If bRL <> 0 And wl.Cells(rr, 5).Value = "" Then
            bRet = bRL / 550000
            wl.Cells(rr, 5).Value = bRet
            wl.Cells(rr, 5).NumberFormat = "0.00%"
            ColorPnl wl.Cells(rr, 5), bRet
        End If
        ' Backfill col 11 (Combined RL PnL) -- col4 + col8 if empty
        If bRL <> 0 And wl.Cells(rr, 11).Value = "" Then
            bComb = bRL + V2D(wl.Cells(rr, 8).Value)
            wl.Cells(rr, 11).Value = bComb
            wl.Cells(rr, 11).NumberFormat = "#,##0;[Red](#,##0)"
            ColorPnl wl.Cells(rr, 11), bComb
        End If
        ' Backfill col 12 (Daily Ret Chg%) -- (this row's Combined PNL - prev row's) / prev row's |Combined MV|
        If rr > 3 And wl.Cells(rr, 12).Value = "" Then
            Dim bPrevMV As Double, bPrevPNL As Double, bDRet As Double
            bPrevMV  = V2D(wl.Cells(rr - 1, 9).Value)
            bPrevPNL = V2D(wl.Cells(rr - 1, 10).Value)
            If bPrevMV <> 0 Then
                bDRet = (V2D(wl.Cells(rr, 10).Value) - bPrevPNL) / Abs(bPrevMV)
                wl.Cells(rr, 12).Value = bDRet
                wl.Cells(rr, 12).NumberFormat = "0.00%;[Red](0.00%)"
                ColorPnl wl.Cells(rr, 12), bDRet
            End If
        End If
    Next rr
    Call BuildRLChart(wl)
    Call BuildRiskMetrics(wl, lr)
    wl.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    Application.StatusBar = "CPLog ready  " & Format(Now, "hh:mm:ss")
End Sub


' =============================================================
' PRIVATE: ensure CPLog sheet -- 14-col schema, colored headers
' Row 1 = title  Row 2 = headers  Row 3+ = data
' =============================================================
Private Function EnsureLogSheet() As Worksheet
    Dim wl As Worksheet
    Dim shIsNew As Boolean: shIsNew = False
    On Error Resume Next: Set wl = ThisWorkbook.Sheets(SH_LOG): On Error GoTo 0

    If wl Is Nothing Then
        shIsNew = True
        Dim ref As Worksheet
        On Error Resume Next: Set ref = ThisWorkbook.Sheets(SH_OUT): On Error GoTo 0
        If ref Is Nothing Then
            Set wl = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        Else
            Set wl = ThisWorkbook.Sheets.Add(After:=ref)
        End If
        wl.Name = SH_LOG
        wl.Cells.Interior.Color = CLR_BG_DARK
        wl.Cells.Font.Color     = CLR_FG_SILVER
        wl.Cells.Font.Size      = 9
    End If

    ' --- ALWAYS: refresh title (row 1) ---
    wl.Rows(1).RowHeight = 24
    On Error Resume Next: wl.Range(wl.Cells(1,1), wl.Cells(1,14)).UnMerge: On Error GoTo 0
    With wl.Range(wl.Cells(1, 1), wl.Cells(1, 14))
        .Merge
        .Value               = "RR4 + RR5  Combined History Log"
        .Font.Bold           = True
        .Font.Size           = 13
        .Font.Color          = CLR_FG_WHITE
        .Interior.Color      = CLR_BG_MID
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With

    ' --- ALWAYS: refresh headers (row 2) ---
    wl.Rows(2).RowHeight = 18
    Dim hdr As Variant
    hdr = Array("Date", "RR4 Net MV", "RR4 Unrl PNL", "RR4 RL PNL", "RR4 RL Ret%", _
                "RR5 Total MV", "RR5 Unrl PNL", "RR5 RL PNL", _
                "Combined MV", "Combined PNL", _
                "Combined RL PnL", "Daily Ret Chg%", "TWII Price", "Notes")
    Dim hclr As Variant
    hclr = Array(CLR_FG_WHITE, _
                 CLR_GRP_RR4, CLR_GRP_RR4, CLR_GRP_RR4, CLR_GRP_RR4, _
                 CLR_GRP_RR5, CLR_GRP_RR5, CLR_GRP_RR5, _
                 CLR_GRP_COMB, CLR_GRP_COMB, _
                 CLR_GRP_COMB, CLR_GRP_COMB, CLR_FG_SILVER, CLR_FG_SILVER)
    Dim c As Long
    For c = 1 To 14
        With wl.Cells(2, c)
            .Value               = hdr(c - 1)
            .Font.Bold           = True
            .Font.Color          = hclr(c - 1)
            .Interior.Color      = CLR_BG_MID
            .HorizontalAlignment = xlCenter
            .VerticalAlignment   = xlCenter
        End With
    Next c

    ' --- ALWAYS: column widths ---
    wl.Columns(1).ColumnWidth  = 11
    wl.Columns(2).ColumnWidth  = 12
    wl.Columns(3).ColumnWidth  = 12
    wl.Columns(4).ColumnWidth  = 12
    wl.Columns(5).ColumnWidth  = 11
    wl.Columns(6).ColumnWidth  = 12
    wl.Columns(7).ColumnWidth  = 12
    wl.Columns(8).ColumnWidth  = 12
    wl.Columns(9).ColumnWidth  = 12
    wl.Columns(10).ColumnWidth = 13
    wl.Columns(11).ColumnWidth = 13
    wl.Columns(12).ColumnWidth = 11
    wl.Columns(13).ColumnWidth = 12
    wl.Columns(14).ColumnWidth = 16

    ' --- ONLY new sheets: seed row at LOG_START ---
    If shIsNew Then
        Dim twdFmt0 As String: twdFmt0 = "#,##0;[Red](#,##0)"
        wl.Cells(3, 1).Value        = LOG_START
        wl.Cells(3, 1).NumberFormat = "yyyy/m/d"
        wl.Cells(3, 2).Value        = 0: wl.Cells(3, 2).NumberFormat = twdFmt0
        wl.Cells(3, 3).Value        = 0: wl.Cells(3, 3).NumberFormat = twdFmt0
        wl.Cells(3, 4).Value        = 0: wl.Cells(3, 4).NumberFormat = twdFmt0
        wl.Cells(3, 5).Value        = 0: wl.Cells(3, 5).NumberFormat = "0.00%"
        wl.Cells(3, 9).Value        = 0: wl.Cells(3, 9).NumberFormat  = twdFmt0
        wl.Cells(3, 10).Value       = 0: wl.Cells(3, 10).NumberFormat = twdFmt0
        wl.Cells(3, 11).Value       = 0: wl.Cells(3, 11).NumberFormat = twdFmt0
        wl.Rows(3).Interior.Color   = CLR_SECTION
    End If

    Set EnsureLogSheet = wl
End Function


' =============================================================
' PRIVATE: append one log row -- reads live values from CombinedPerf ws
' =============================================================
Private Sub AppendLog(ws As Worksheet)
    On Error Resume Next
    Dim wl As Worksheet: Set wl = EnsureLogSheet()
    If wl Is Nothing Then Exit Sub

    ' --- find existing row for today or append new ---
    Dim nr As Long: nr = 0
    Dim lastR As Long: lastR = wl.Cells(wl.Rows.Count, 1).End(xlUp).Row
    Dim chkR As Long
    For chkR = 3 To lastR
        If Int(wl.Cells(chkR, 1).Value) = Int(Date) Then
            nr = chkR
            Exit For
        End If
    Next chkR
    If nr = 0 Then
        nr = lastR + 1
        If nr < 3 Then nr = 3
    End If

    ' --- read from CombinedPerf sheet ---
    Dim rr4MV   As Double: rr4MV   = SN(ws.Cells(R_NEXPOS,   C_RR4))
    Dim rr4Unrl As Double: rr4Unrl = SN(ws.Cells(R_UNRL_PNL, C_RR4))
    Dim rr4RL   As Double: rr4RL   = SN(ws.Cells(R_RL_PNL,   C_RR4))
    Dim rr4Ret  As Double: rr4Ret  = SN(ws.Cells(R_UNRL_PCT,  C_RR4))
    Dim rr5MV   As Double: rr5MV   = SN(ws.Cells(R_NEXPOS,   C_RR5))
    Dim rr5Unrl As Double: rr5Unrl = SN(ws.Cells(R_UNRL_PNL, C_RR5))
    Dim rr5RL   As Double: rr5RL   = SN(ws.Cells(R_RL_PNL,   C_RR5))
    Dim combMV  As Double: combMV  = SN(ws.Cells(R_NEXPOS,   C_TOTAL))
    Dim combPNL As Double: combPNL = SN(ws.Cells(R_TOT_PNL,  C_TOTAL))

    Dim combRL  As Double: combRL  = rr4RL + rr5RL

    ' --- read from RR5 sheet (TWII) ---
    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Sheets(SH_RR5)
    Dim twiiPx   As Double: twiiPx   = 0
    If Not wsR Is Nothing Then
        twiiPx   = V2D(wsR.Range(RR5_TWII_CELL).Value)
    End If

    Dim twdFmt As String: twdFmt = "#,##0;[Red](#,##0)"

    ' col 1: Date
    wl.Cells(nr, 1).Value        = Date
    wl.Cells(nr, 1).NumberFormat = "yyyy/m/d"
    wl.Cells(nr, 1).Interior.Color = CLR_BG_MID

    ' cols 2-5: RR4 (skip if RR4 not connected)
    If rr4MV <> 0 Then
        wl.Cells(nr, 2).Value        = rr4MV:   wl.Cells(nr, 2).NumberFormat = twdFmt
        wl.Cells(nr, 3).Value        = rr4Unrl: wl.Cells(nr, 3).NumberFormat = twdFmt
        wl.Cells(nr, 4).Value        = rr4RL:   wl.Cells(nr, 4).NumberFormat = twdFmt
        wl.Cells(nr, 5).Value        = rr4Ret:  wl.Cells(nr, 5).NumberFormat = "0.00%"
        ColorPnl wl.Cells(nr, 3), rr4Unrl
        ColorPnl wl.Cells(nr, 4), rr4RL
        ColorPnl wl.Cells(nr, 5), rr4Ret
    End If

    ' cols 6-8: RR5 (always write -- RR5 sheet lives in this workbook, unlike
    ' RR4 which may genuinely not be open. Net exposure can legitimately be 0
    ' while Unrl/RL PnL are not, so gating on rr5MV<>0 dropped real data.)
    wl.Cells(nr, 6).Value        = rr5MV:   wl.Cells(nr, 6).NumberFormat = twdFmt
    wl.Cells(nr, 7).Value        = rr5Unrl: wl.Cells(nr, 7).NumberFormat = twdFmt
    wl.Cells(nr, 8).Value        = rr5RL:   wl.Cells(nr, 8).NumberFormat = twdFmt
    ColorPnl wl.Cells(nr, 7), rr5Unrl
    ColorPnl wl.Cells(nr, 8), rr5RL

    ' cols 9-10: Combined (always)
    wl.Cells(nr, 9).Value         = combMV:  wl.Cells(nr, 9).NumberFormat  = twdFmt
    wl.Cells(nr, 10).Value        = combPNL: wl.Cells(nr, 10).NumberFormat = twdFmt
    ColorPnl wl.Cells(nr, 10), combPNL

    ' col 11: Combined RL PnL
    wl.Cells(nr, 11).Value        = combRL: wl.Cells(nr, 11).NumberFormat = twdFmt: ColorPnl wl.Cells(nr, 11), combRL

    ' col 12: Daily Ret Chg% = (today's Combined PNL - prev row's) / prev row's |Combined MV|
    If nr > 3 Then
        Dim prevMV As Double, prevPNL As Double, dRet As Double
        prevMV  = V2D(wl.Cells(nr - 1, 9).Value)
        prevPNL = V2D(wl.Cells(nr - 1, 10).Value)
        If prevMV <> 0 Then
            dRet = (combPNL - prevPNL) / Abs(prevMV)
            wl.Cells(nr, 12).Value        = dRet
            wl.Cells(nr, 12).NumberFormat = "0.00%;[Red](0.00%)"
            ColorPnl wl.Cells(nr, 12), dRet
        End If
    End If
    ' col 13: TWII Price
    If twiiPx <> 0 Then
        wl.Cells(nr, 13).Value        = twiiPx
        wl.Cells(nr, 13).NumberFormat = "#,##0.00"
    End If

    ' alternating gray / dark row colors
    wl.Rows(nr).Interior.Color = IIf(nr Mod 2 = 1, CLR_SECTION, CLR_BG_MID)
    On Error GoTo 0
End Sub


' =============================================================
' PUBLIC: seed one historical log row (call from VBA Immediate window)
' e.g.: SeedLogRow #1/1/2026#, 800000, -20000, 350000, 0.636, 0,0,0, 0,0,0
' =============================================================
Public Sub SeedLogRow(dt As Date, _
    r4MV As Double, r4Up As Double, r4RP As Double, r4Ret As Double, _
    r5MV As Double, r5Up As Double, r5RP As Double, _
    combRL As Double, ivA As Double, twii As Double)
    On Error Resume Next
    Dim wl As Worksheet: Set wl = EnsureLogSheet()
    If wl Is Nothing Then Exit Sub

    Dim nr As Long
    nr = wl.Cells(wl.Rows.Count, 1).End(xlUp).Row + 1
    If nr < 3 Then nr = 3

    Dim twdFmt As String: twdFmt = "#,##0;[Red](#,##0)"

    wl.Cells(nr, 1).Value        = dt
    wl.Cells(nr, 1).NumberFormat = "yyyy/m/d"
    wl.Cells(nr, 1).Interior.Color = CLR_BG_MID

    If r4MV <> 0 Then
        wl.Cells(nr, 2).Value = r4MV:  wl.Cells(nr, 2).NumberFormat = twdFmt
        wl.Cells(nr, 3).Value = r4Up:  wl.Cells(nr, 3).NumberFormat = twdFmt:  ColorPnl wl.Cells(nr, 3), r4Up
        wl.Cells(nr, 4).Value = r4RP:  wl.Cells(nr, 4).NumberFormat = twdFmt:  ColorPnl wl.Cells(nr, 4), r4RP
        wl.Cells(nr, 5).Value = r4Ret: wl.Cells(nr, 5).NumberFormat = "0.00%": ColorPnl wl.Cells(nr, 5), r4Ret
    End If
    If r5MV <> 0 Then
        wl.Cells(nr, 6).Value = r5MV:  wl.Cells(nr, 6).NumberFormat = twdFmt
        wl.Cells(nr, 7).Value = r5Up:  wl.Cells(nr, 7).NumberFormat = twdFmt:  ColorPnl wl.Cells(nr, 7), r5Up
        wl.Cells(nr, 8).Value = r5RP:  wl.Cells(nr, 8).NumberFormat = twdFmt:  ColorPnl wl.Cells(nr, 8), r5RP
    End If

    Dim combMV  As Double: combMV  = r4MV + r5MV
    Dim combPNL As Double: combPNL = (r4Up + r4RP) + (r5Up + r5RP)
    wl.Cells(nr, 9).Value        = combMV:  wl.Cells(nr, 9).NumberFormat  = twdFmt
    wl.Cells(nr, 10).Value       = combPNL: wl.Cells(nr, 10).NumberFormat = twdFmt: ColorPnl wl.Cells(nr, 10), combPNL

    If combRL <> 0 Then wl.Cells(nr, 11).Value = combRL: wl.Cells(nr, 11).NumberFormat = "#,##0;[Red](#,##0)": ColorPnl wl.Cells(nr, 11), combRL
    If ivA    <> 0 Then wl.Cells(nr, 12).Value = ivA:    wl.Cells(nr, 12).NumberFormat = "0.00%"
    If twii   <> 0 Then wl.Cells(nr, 13).Value = twii:   wl.Cells(nr, 13).NumberFormat = "#,##0.00"

    wl.Rows(nr).Interior.Color = IIf(nr Mod 2 = 1, CLR_SECTION, CLR_BG_MID)
    On Error GoTo 0
End Sub


' =============================================================
' PUBLIC: import RR4 HistoryLog -> CPLog  (run once, RR4 wb must be open)
' RR4 HistoryLog columns: A=DateTime  B=NetMV  C=TotalPNL  F=TWII  G=RealizedPNL
' Derives:  Unrl PNL = C - G   |  Combined MV = B  |  Combined PNL = C
' Skips rows already having RR4 data.  Sorts CPLog by date when done.
' =============================================================
Public Sub ImportRR4History()
    Dim wl As Worksheet: Set wl = EnsureLogSheet()
    If wl Is Nothing Then Exit Sub

    Dim r4wb As Workbook: Set r4wb = FindRR4WB()
    If r4wb Is Nothing Then
        MsgBox "RR4 workbook not open.  Open it first.", vbExclamation, "ImportRR4History"
        Exit Sub
    End If

    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = r4wb.Sheets("HistoryLog"): On Error GoTo 0
    If wsH Is Nothing Then
        MsgBox "HistoryLog sheet not found in RR4 workbook.", vbExclamation, "ImportRR4History"
        Exit Sub
    End If

    Dim hLast As Long: hLast = wsH.Cells(wsH.Rows.Count, "B").End(xlUp).Row
    If hLast < 2 Then MsgBox "RR4 HistoryLog is empty.", vbInformation, "ImportRR4History": Exit Sub

    Dim cpLast As Long: cpLast = wl.Cells(wl.Rows.Count, 1).End(xlUp).Row
    Dim maxRows As Long: maxRows = hLast + cpLast + 10

    ' existing date -> row map
    Dim exDate() As Date:  ReDim exDate(1 To maxRows)
    Dim exRow()  As Long:  ReDim exRow(1 To maxRows)
    Dim exCnt    As Long:  exCnt = 0
    Dim cr As Long
    For cr = 3 To cpLast
        Dim cdv As Variant: cdv = wl.Cells(cr, 1).Value
        If IsDate(cdv) Then
            exCnt = exCnt + 1
            exDate(exCnt) = Int(CDate(cdv))
            exRow(exCnt)  = cr
        End If
    Next cr

    Dim twdFmt As String: twdFmt = "#,##0;[Red](#,##0)"
    Dim added As Long: added = 0

    Dim hr As Long
    For hr = 2 To hLast
        ' --- read date ---
        Dim rawDt As Variant: rawDt = wsH.Cells(hr, "A").Value
        If Not IsDate(rawDt) Then GoTo SkipHR
        Dim dt As Date: dt = Int(CDate(rawDt))
        If dt < LOG_START Then GoTo SkipHR

        ' --- read RR4 values ---
        Dim r4MV   As Double: r4MV   = V2D(wsH.Cells(hr, "B").Value)
        Dim r4Tot  As Double: r4Tot  = V2D(wsH.Cells(hr, "C").Value)
        Dim r4RL   As Double: r4RL   = V2D(wsH.Cells(hr, "G").Value)
        Dim r4Unrl As Double: r4Unrl = r4Tot - r4RL
        Dim twii   As Double: twii   = V2D(wsH.Cells(hr, "F").Value)

        ' --- find or create CPLog row ---
        Dim wr As Long: wr = 0
        Dim di As Long
        For di = 1 To exCnt
            If exDate(di) = dt Then wr = exRow(di): Exit For
        Next di

        If wr > 0 Then
            ' Always update col 11 (Combined RL PnL) even for existing rows
            wl.Cells(wr, 11).Value = r4RL: wl.Cells(wr, 11).NumberFormat = twdFmt: ColorPnl wl.Cells(wr, 11), r4RL
            ' Always update col 5 (RR4 RL Ret% = Realized PNL / 550k)
            wl.Cells(wr, 5).Value = r4RL / 550000: wl.Cells(wr, 5).NumberFormat = "0.00%": ColorPnl wl.Cells(wr, 5), r4RL / 550000
            ' Skip other fields if row already has RR4 data
            If V2D(wl.Cells(wr, 2).Value) <> 0 Then GoTo SkipHR
        Else
            ' new row -- append
            cpLast = cpLast + 1
            wr = cpLast
            wl.Rows(wr).Interior.Color = IIf(wr Mod 2 = 1, CLR_SECTION, CLR_BG_MID)
            exCnt = exCnt + 1
            exDate(exCnt) = dt
            exRow(exCnt)  = wr
            added = added + 1
        End If

        ' --- write to CPLog ---
        wl.Cells(wr, 1).Value        = dt
        wl.Cells(wr, 1).NumberFormat = "yyyy/m/d"
        wl.Cells(wr, 1).Interior.Color = CLR_BG_MID

        wl.Cells(wr, 2).Value        = r4MV:   wl.Cells(wr, 2).NumberFormat = twdFmt
        wl.Cells(wr, 3).Value        = r4Unrl: wl.Cells(wr, 3).NumberFormat = twdFmt: ColorPnl wl.Cells(wr, 3), r4Unrl
        wl.Cells(wr, 4).Value        = r4RL:   wl.Cells(wr, 4).NumberFormat = twdFmt: ColorPnl wl.Cells(wr, 4), r4RL
        ' col 5: RR4 RL Ret% = Realized PNL / 550k capital base
        wl.Cells(wr, 5).Value = r4RL / 550000: wl.Cells(wr, 5).NumberFormat = "0.00%": ColorPnl wl.Cells(wr, 5), r4RL / 550000
        ' col 11: Combined RL PnL = RR4 RL (RR5 not yet running for historical data)
        wl.Cells(wr, 11).Value = r4RL: wl.Cells(wr, 11).NumberFormat = twdFmt: ColorPnl wl.Cells(wr, 11), r4RL

        ' Combined (RR4-only for history rows where RR5 not yet running)
        wl.Cells(wr, 9).Value        = r4MV:   wl.Cells(wr, 9).NumberFormat  = twdFmt
        ' Combined PNL only if RR5 RL col is blank
        If V2D(wl.Cells(wr, 8).Value) = 0 Then
            wl.Cells(wr, 10).Value   = r4Tot:  wl.Cells(wr, 10).NumberFormat = twdFmt: ColorPnl wl.Cells(wr, 10), r4Tot
        End If

        If twii <> 0 Then
            wl.Cells(wr, 13).Value        = twii
            wl.Cells(wr, 13).NumberFormat = "#,##0.00"
        End If

SkipHR:
    Next hr

    ' sort CPLog data rows by date ascending
    Dim cpFinal As Long: cpFinal = wl.Cells(wl.Rows.Count, 1).End(xlUp).Row
    If cpFinal > 3 Then
        wl.Range(wl.Cells(3, 1), wl.Cells(cpFinal, 14)).Sort _
            Key1:=wl.Columns(1), Order1:=xlAscending, Header:=xlNo
    End If
    ' Re-stamp alternating row colors after sort scrambles them
    Dim srr As Long
    For srr = 3 To cpFinal
        wl.Rows(srr).Interior.Color = IIf(srr Mod 2 = 1, CLR_SECTION, CLR_BG_MID)
    Next srr

    MsgBox added & " RR4 history rows added.  Click [Refresh All] (or run BuildCombinedHistoryLog) to rebuild the chart.", _
           vbInformation, "ImportRR4History"
End Sub


' =============================================================
' PRIVATE: build / refresh RL PNL line chart on CPLog sheet
' =============================================================
Private Sub BuildRLChart(wl As Worksheet)
    ' --- 1. Delete any existing CPLogLine_x charts ---
    On Error Resume Next
    Dim co As ChartObject
    For Each co In wl.ChartObjects
        If Left(co.Name, 10) = "CPLogLine_" Then co.Delete
    Next co
    On Error GoTo 0

    ' --- 2. Determine data extent (col A = dates, rows 3..lastRow) ---
    Dim lastRow As Long
    lastRow = wl.Cells(wl.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    Dim xRng As Range
    Set xRng = wl.Range(wl.Cells(3, 1), wl.Cells(lastRow, 1))

    ' X axis floor = first data date (Excel serial); prevents axis drifting to 2025
    Dim xMin As Double
    If IsDate(wl.Cells(3, 1).Value) Then
        xMin = CDbl(CDate(wl.Cells(3, 1).Value))
    Else
        xMin = CDbl(CDate(LOG_START))
    End If

    ' --- 3. Series definition: source col, title, line color, Y number format ---
    Dim sCols(0 To 7) As Long
    Dim sNams(0 To 7) As String
    Dim sClrs(0 To 7) As Long
    Dim sFmts(0 To 7) As String
    sCols(0) = 3:  sNams(0) = "RR4 Unrl PNL":  sClrs(0) = CLR_GRP_RR4:   sFmts(0) = "#,##0"
    sCols(1) = 4:  sNams(1) = "RR4 RL PNL":    sClrs(1) = CLR_GRP_RR4:   sFmts(1) = "#,##0"
    sCols(2) = 5:  sNams(2) = "RR4 RL Ret%":   sClrs(2) = CLR_GRP_RR4:   sFmts(2) = "0.00%"
    sCols(3) = 6:  sNams(3) = "RR5 Total MV":  sClrs(3) = CLR_GRP_RR5:   sFmts(3) = "#,##0"
    sCols(4) = 7:  sNams(4) = "RR5 Unrl PNL":  sClrs(4) = CLR_GRP_RR5:   sFmts(4) = "#,##0"
    sCols(5) = 8:  sNams(5) = "RR5 RL PNL":    sClrs(5) = CLR_GRP_RR5:   sFmts(5) = "#,##0"
    sCols(6) = 9:  sNams(6) = "Combined MV":   sClrs(6) = CLR_GRP_COMB:  sFmts(6) = "#,##0"
    sCols(7) = 10: sNams(7) = "Combined PNL":  sClrs(7) = CLR_GRP_COMB:  sFmts(7) = "#,##0"

    Dim bgClr As Long:   bgClr = RGB(27, 27, 27)     ' #1B1B1B chart background
    Dim gridClr As Long: gridClr = RGB(55, 55, 55)   ' subtle horizontal gridlines

    ' --- 4. Build one chart per series (single column, stacked vertically) ---
    Dim si As Long
    For si = 0 To 7
        Dim sRow As Long: sRow = 3 + si * 29   ' row 3 start; 28 rows chart + 1 row gap
        Dim sCol As Long: sCol = 15            ' col O

        Dim cLeft   As Double: cLeft   = wl.Cells(sRow, sCol).Left
        Dim cTop    As Double: cTop    = wl.Cells(sRow, 1).Top
        Dim cWidth  As Double: cWidth  = wl.Cells(sRow, sCol + 14).Left - cLeft
        Dim cHeight As Double: cHeight = wl.Cells(sRow + 28, 1).Top - cTop

        Dim co2 As ChartObject
        Set co2 = wl.ChartObjects.Add(cLeft, cTop, cWidth, cHeight)
        co2.Name = "CPLogLine_" & si

        With co2.Chart
            ' Line chart: category axis can be a TRUE time scale (calendar months)
            .ChartType = 4  ' xlLine

            ' Wipe any auto-created series, then add exactly one
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            With .SeriesCollection.NewSeries
                .Values  = wl.Range(wl.Cells(3, sCols(si)), wl.Cells(lastRow, sCols(si)))
                .XValues = xRng
                .Name    = sNams(si)
                .Border.Color  = sClrs(si)
                .Border.Weight = 2   ' xlThin -- clearly visible
                .Smooth = False
            End With

            ' Fonts AFTER ChartType so Excel does not reset them
            .ChartArea.Font.Name  = "Consolas"
            .ChartArea.Font.Color = CLR_FG_WHITE
            .ChartArea.Font.Size  = 9

            ' Backgrounds
            .PlotArea.Interior.Color   = bgClr
            .PlotArea.Border.LineStyle = -4142
            .ChartArea.Interior.Color  = bgClr

            ' Outer border matches this series' group color (RR4/RR5/Combined)
            With .ChartArea.Border
                .LineStyle = 1
                .Color     = sClrs(si)
                .Weight    = 2
            End With

            ' --- X axis: TIME SCALE = one tick per calendar month ---
            On Error Resume Next
            With .Axes(1)   ' xlCategory
                .CategoryType       = 3   ' xlTimeScale -- treat X values as real dates
                .BaseUnit           = 0   ' xlDays  -- day granularity underneath
                .BaseUnitIsAuto     = False
                .MajorUnitScale     = 1   ' xlMonths -- MAJOR TICK EVERY CALENDAR MONTH
                .MajorUnit          = 1
                .MajorUnitIsAuto    = False
                .MinimumScaleIsAuto = False
                .MinimumScale       = xMin   ' left edge = first data date
                .MaximumScaleIsAuto = True   ' right edge follows the latest date
                .TickLabels.NumberFormat = "yyyy/m"
                .TickLabels.Font.Name    = "Consolas"
                .TickLabels.Font.Color   = CLR_FG_WHITE
                .TickLabels.Font.Size    = 8
                .AxisLine.Border.LineStyle = -4142
                .MajorGridlines.Delete
            End With
            On Error GoTo 0

            ' --- Y axis ---
            On Error Resume Next
            With .Axes(2)   ' xlValue
                .TickLabels.NumberFormat   = sFmts(si)
                .TickLabels.Font.Name      = "Consolas"
                .TickLabels.Font.Color     = CLR_FG_WHITE
                .TickLabels.Font.Size      = 8
                .AxisLine.Border.LineStyle = -4142
                .MajorGridlines.Border.Color  = gridClr
                .MajorGridlines.Border.Weight = 1
            End With
            On Error GoTo 0

            ' Title
            .HasTitle = True
            .ChartTitle.Text       = sNams(si)
            .ChartTitle.Font.Name  = "Consolas"
            .ChartTitle.Font.Color = CLR_FG_WHITE
            .ChartTitle.Font.Bold  = False
            .ChartTitle.Font.Size  = 10

            .HasLegend = False
        End With
        Set co2 = Nothing
    Next si
End Sub


' =============================================================
' PRIVATE: risk-adjusted performance panel (CPLog cols AD:AF)
' 30D volatility + Sharpe-like ratio (last 30 rows of col 12),
' Max Drawdown (compounds col 12 into a cumulative index),
' Beta / Correlation vs TWII (col 13 daily change vs col 12).
' Placed clear of the 8 stacked charts (which occupy cols O:AB).
' =============================================================
Private Sub BuildRiskMetrics(wl As Worksheet, lastRow As Long)
    Dim mc As Long: mc = 30   ' AD -- metric label
    Dim vc As Long: vc = 31   ' AE -- value
    Dim dc As Long: dc = 32   ' AF -- detail

    wl.Range(wl.Cells(3, mc), wl.Cells(9, dc)).Clear
    wl.Columns(mc).ColumnWidth = 20
    wl.Columns(vc).ColumnWidth = 10
    wl.Columns(dc).ColumnWidth = 24

    If lastRow < 5 Then Exit Sub   ' need a few days of history before this means anything

    ' --- 30D volatility + Sharpe-like ratio: last 30 valid col-12 values ---
    Dim volArr() As Double: ReDim volArr(1 To 30)
    Dim volN As Long: volN = 0
    Dim rr As Long
    For rr = lastRow To 4 Step -1
        If volN >= 30 Then Exit For
        If IsNumeric(wl.Cells(rr, 12).Value) And wl.Cells(rr, 12).Value <> "" Then
            volN = volN + 1
            volArr(volN) = CDbl(wl.Cells(rr, 12).Value)
        End If
    Next rr

    Dim dailyVol As Double, annVol As Double, dailyMean As Double, sharpe As Double
    If volN >= 2 Then
        Dim volSlice() As Double: ReDim volSlice(1 To volN)
        Dim vi As Long
        For vi = 1 To volN: volSlice(vi) = volArr(vi): Next vi
        dailyVol  = Application.WorksheetFunction.StDev(volSlice)
        dailyMean = Application.WorksheetFunction.Average(volSlice)
        annVol    = dailyVol * Sqr(252)
        If dailyVol <> 0 Then sharpe = (dailyMean / dailyVol) * Sqr(252)
    End If

    ' --- Max Drawdown: compound col 12 into a cumulative index, track running peak ---
    Dim cumIdx As Double: cumIdx = 1
    Dim peakIdx As Double: peakIdx = 1
    Dim peakDateRun As Date, ddPeakDate As Date, ddTroughDate As Date
    Dim maxDD As Double: maxDD = 0
    Dim firstDD As Boolean: firstDD = True
    Dim ddRow As Long
    For ddRow = 4 To lastRow
        If IsNumeric(wl.Cells(ddRow, 12).Value) And wl.Cells(ddRow, 12).Value <> "" Then
            cumIdx = cumIdx * (1 + CDbl(wl.Cells(ddRow, 12).Value))
            If firstDD Then
                peakDateRun = CDate(wl.Cells(ddRow, 1).Value)
                firstDD = False
            End If
            If cumIdx > peakIdx Then
                peakIdx = cumIdx
                peakDateRun = CDate(wl.Cells(ddRow, 1).Value)
            End If
            Dim curDD As Double: curDD = (cumIdx - peakIdx) / peakIdx
            If curDD < maxDD Then
                maxDD = curDD
                ddPeakDate = peakDateRun
                ddTroughDate = CDate(wl.Cells(ddRow, 1).Value)
            End If
        End If
    Next ddRow

    ' --- Beta / Correlation vs TWII: pair col 12 (port ret) with TWII's own daily change ---
    Dim portArr() As Double: ReDim portArr(1 To lastRow)
    Dim twiiArr() As Double: ReDim twiiArr(1 To lastRow)
    Dim betaN As Long: betaN = 0
    Dim br As Long
    For br = 4 To lastRow
        Dim tNow As Variant, tPrev As Variant, pRet As Variant
        tNow = wl.Cells(br, 13).Value
        tPrev = wl.Cells(br - 1, 13).Value
        pRet = wl.Cells(br, 12).Value
        If IsNumeric(tNow) And IsNumeric(tPrev) And CDbl(tPrev) <> 0 And IsNumeric(pRet) And pRet <> "" Then
            betaN = betaN + 1
            portArr(betaN) = CDbl(pRet)
            twiiArr(betaN) = (CDbl(tNow) - CDbl(tPrev)) / CDbl(tPrev)
        End If
    Next br

    Dim beta As Double, corr As Double
    Dim haveBeta As Boolean: haveBeta = False
    If betaN >= 5 Then
        Dim pSlice() As Double: ReDim pSlice(1 To betaN)
        Dim tSlice() As Double: ReDim tSlice(1 To betaN)
        Dim bi As Long
        For bi = 1 To betaN: pSlice(bi) = portArr(bi): tSlice(bi) = twiiArr(bi): Next bi
        On Error Resume Next
        Err.Clear
        beta = Application.WorksheetFunction.Slope(pSlice, tSlice)
        corr = Application.WorksheetFunction.Correl(pSlice, tSlice)
        If Err.Number = 0 Then haveBeta = True
        On Error GoTo 0
    End If

    ' --- Render: title bar + header row (base styling only, no colors yet) ---
    With wl.Range(wl.Cells(3, mc), wl.Cells(3, dc))
        .Merge
        .Value = "RISK-ADJUSTED PERFORMANCE"
        .Font.Name = "Consolas": .Font.Bold = True: .Font.Size = 11
        .Font.Color = CLR_FG_WHITE: .Interior.Color = CLR_BG_MID
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    wl.Rows(3).RowHeight = 20

    Dim hr As Long: hr = 4
    With wl.Range(wl.Cells(hr, mc), wl.Cells(hr, dc))
        .Interior.Color = CLR_BG_MID
        .Font.Name = "Consolas": .Font.Bold = True: .Font.Size = 9
        .HorizontalAlignment = xlCenter
    End With
    wl.Cells(hr, mc).Value = "Metric"
    wl.Cells(hr, vc).Value = "Value"
    wl.Cells(hr, dc).Value = "Detail"

    Dim pr As Long: pr = 5

    wl.Cells(pr, mc).Value = "30D Volatility"
    If volN >= 2 Then
        wl.Cells(pr, vc).Value = dailyVol: wl.Cells(pr, vc).NumberFormat = "0.00%"
        wl.Cells(pr, dc).Value = "annualized " & Format(annVol, "0%")
    Else
        wl.Cells(pr, vc).Value = "-":  wl.Cells(pr, dc).Value = "need more history"
    End If
    pr = pr + 1

    wl.Cells(pr, mc).Value = "Sharpe-like"
    If volN >= 2 And dailyVol <> 0 Then
        wl.Cells(pr, vc).Value = sharpe: wl.Cells(pr, vc).NumberFormat = "0.00"
        wl.Cells(pr, dc).Value = "return / volatility"
    Else
        wl.Cells(pr, vc).Value = "-":  wl.Cells(pr, dc).Value = "need more history"
    End If
    pr = pr + 1

    wl.Cells(pr, mc).Value = "Max Drawdown"
    If maxDD < 0 Then
        wl.Cells(pr, vc).Value = maxDD: wl.Cells(pr, vc).NumberFormat = "0.00%"
        wl.Cells(pr, dc).Value = Format(ddPeakDate, "yyyy/m/d") & " to " & Format(ddTroughDate, "yyyy/m/d")
    Else
        wl.Cells(pr, vc).Value = "-":  wl.Cells(pr, dc).Value = "no drawdown yet"
    End If
    pr = pr + 1

    wl.Cells(pr, mc).Value = "Beta vs TWII"
    If haveBeta Then
        wl.Cells(pr, vc).Value = beta: wl.Cells(pr, vc).NumberFormat = "0.00"
        wl.Cells(pr, dc).Value = "corr " & Format(corr, "0.00")
    Else
        wl.Cells(pr, vc).Value = "-":  wl.Cells(pr, dc).Value = "TWII data too sparse"
    End If

    Dim lastPr As Long: lastPr = pr

    ' --- Base styling for the whole data block (BEFORE per-value colors below) ---
    With wl.Range(wl.Cells(5, mc), wl.Cells(lastPr, dc))
        .Font.Name = "Consolas": .Font.Size = 9: .Font.Color = CLR_FG_WHITE
        .Borders.LineStyle = xlContinuous: .Borders.Color = CLR_BORDER
    End With
    Dim rz As Long
    For rz = 5 To lastPr
        wl.Rows(rz).Interior.Color = IIf(rz Mod 2 = 1, CLR_SECTION, CLR_BG_MID)
    Next rz

    ' --- Per-value colors applied LAST so the block styling above doesn't clobber them ---
    If volN >= 2 And dailyVol <> 0 Then ColorPnl wl.Cells(6, vc), sharpe
    If maxDD < 0 Then ColorPnl wl.Cells(7, vc), maxDD
End Sub


' =============================================================
' PUBLIC: Build holdings pie chart (RR5 breakdown by type)
' Reads the Futures/Options/Warrants net-exposure rows from
' the CombinedPerf sheet and renders a pie chart below the table.
' =============================================================
Public Sub BuildHoldingsPieCharts()
    ' Builds TWO pie charts on CombinedPerf, stacked vertically under E4:
    '   1. "MARKET VALUE DISTRIBUTION"    -- raw market value, unadjusted
    '   2. "LEVERAGE-ADJUSTED EXPOSURE"   -- weighted by leverage (see below)
    ' Both are kept so raw capital-at-risk and true economic exposure can be
    ' compared side by side -- neither replaces the other.
    ' Data sources (raw vs leveraged):
    '   RR5 Futures  rows 7-14   col N (14) = NET EXPOS (already full notional, same in both)
    '   RR5 Options  rows 19-26  col Q (17) = NET EXPOS (same in both)
    '   RR5 Warrants rows 31-38  col Q (17) = NET EXPOS; leveraged = NET EXPOS x col AB (28) EFF GEARING
    '   RR4 Analysis col A = Ticker, col G (RR4_PIE_VAL_COL) = NET EXPOS;
    '                        leveraged = NET EXPOS x GetLeverageMultiplier (2x for "NNNNNL" ETFs, etc.)
    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' --- 1. Collect individual positions into two dicts: raw MV and leveraged MV ---
    Dim dictRaw As Object: Set dictRaw = CreateObject("Scripting.Dictionary")
    Dim dictLev As Object: Set dictLev = CreateObject("Scripting.Dictionary")

    ' RR5 ------------------------------------------------------------------
    Dim wsR As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets("RR5")
    On Error GoTo Fail

    If Not wsR Is Nothing Then
        Dim i As Long, code As String, mv As Double

        ' Futures  rows 7-14  col N=14 -- no leverage concept here, same in both dicts
        For i = 7 To 14
            code = Trim(CStr(wsR.Cells(i, 1).Value))
            If code = "" Then GoTo NextFut
            mv = Abs(V2D(wsR.Cells(i, 14).Value))
            If mv > 0 Then
                PieAddDict dictRaw, code, mv
                PieAddDict dictLev, code, mv
            End If
NextFut:
        Next i

        ' Options  rows 19-26  col Q=17 -- same in both dicts
        For i = 19 To 26
            code = Trim(CStr(wsR.Cells(i, 1).Value))
            If code = "" Then GoTo NextOpt
            mv = Abs(V2D(wsR.Cells(i, 17).Value))
            If mv > 0 Then
                PieAddDict dictRaw, code, mv
                PieAddDict dictLev, code, mv
            End If
NextOpt:
        Next i

        ' Warrants rows 31-38  col Q=17 NET EXPOS (capital deployed)
        ' leveraged = x col AB=28 EFF GEARING -> underlying-equivalent exposure
        For i = 31 To 38
            code = Trim(CStr(wsR.Cells(i, 1).Value))
            If code = "" Then GoTo NextWrt
            mv = Abs(V2D(wsR.Cells(i, 17).Value))
            If mv > 0 Then
                PieAddDict dictRaw, code, mv
                Dim wGear As Double: wGear = V2D(wsR.Cells(i, 28).Value)
                If wGear <= 0 Then wGear = 1   ' EFF GEARING not yet computed -- fall back to raw MV
                PieAddDict dictLev, code, mv * wGear
            End If
NextWrt:
        Next i
    End If

    ' RR4 ------------------------------------------------------------------
    Dim rr4wb As Workbook: Set rr4wb = FindRR4WB()
    If Not rr4wb Is Nothing Then
        Dim wsA As Worksheet
        ' col G = market value in RR4 sheet (user confirmed)
        On Error Resume Next: Set wsA = rr4wb.Sheets("RR4"): On Error GoTo Fail
        If Not wsA Is Nothing Then
            ' RR4 sheet layout (DrawColumnHeaders):
            '   Row 4 headers: ISIN, NAME, ENTRY DT, DAYS, BROKER, SECTOR,
            '                  NET EXPOS(7), SHARES, ENTRY PX, LAST, % CHG,
            '                  UNRL PNL, WT%, W.BETA, Beta 180D, P.TARGET, SWING RISK(17)
            '   Data rows start at row 5; broker header / subtotal rows also present.
            Dim r As Long
            Dim valCol As Long: valCol = 7   ' NET EXPOS = col G

            For r = 5 To 500
                Dim tk As String: tk = Trim(CStr(wsA.Cells(r, 1).Value))
                If tk = "" Then GoTo NextRR4
                ' Skip broker header rows ("== BROKER ===") and subtotal rows
                If InStr(tk, "=") > 0 Or InStr(LCase(tk), "subtotal") > 0 Then GoTo NextRR4
                Dim mvRaw As Variant: mvRaw = wsA.Cells(r, valCol).Value
                If Not IsNumeric(mvRaw) Then GoTo NextRR4
                Dim mvAbs As Double: mvAbs = Abs(CDbl(mvRaw))
                If mvAbs > 0 Then
                    PieAddDict dictRaw, tk, mvAbs
                    PieAddDict dictLev, tk, mvAbs * GetLeverageMultiplier(tk)
                End If
NextRR4:
            Next r
        End If
    End If

    If dictRaw.Count = 0 Then
        MsgBox "No position data found." & vbCrLf & _
               "Refresh RR5 first, and open the RR4 workbook if applicable.", _
               vbInformation, "CombinedPerf"
        GoTo Done
    End If

    ' --- 2. Build both pie charts on CombinedPerf sheet, stacked vertically ---
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(SH_OUT): On Error GoTo Fail
    If ws Is Nothing Then Set ws = ActiveSheet

    Dim anchorLeft As Double: anchorLeft = ws.Cells(4, 5).Left
    Dim anchorTop  As Double: anchorTop  = ws.Cells(4, 5).Top

    ' Same ticker = same color in both pies, regardless of each pie's own
    ' value-sorted slice order (alphabetical key -> stable day to day too).
    Dim colorMap As Object: Set colorMap = BuildTickerColorMap(dictRaw, dictLev)

    Call RenderHoldingsPie(ws, dictRaw, "PieHoldings", _
        "RR4 + RR5  MARKET VALUE DISTRIBUTION", "Market Value", _
        anchorLeft, anchorTop, colorMap)

    Call RenderHoldingsPie(ws, dictLev, "PieHoldingsLev", _
        "RR4 + RR5  LEVERAGE-ADJUSTED EXPOSURE", "Leveraged Exposure", _
        anchorLeft, anchorTop + 340, colorMap)

Done:
    Application.ScreenUpdating = True
    Application.StatusBar = "Holdings pie charts built  " & Format(Now, "hh:mm:ss")
    If Not ws Is Nothing Then ws.Activate
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "BuildHoldingsPieCharts failed: " & Err.Description, vbExclamation, "CombinedPerf"
End Sub

' Builds a ticker -> RGB color map from the union of both dicts' keys.
' Assignment order is ALPHABETICAL (not by value/rank), so a given ticker's
' color stays fixed day to day even as market values and rankings shift.
Private Function BuildTickerColorMap(dict1 As Object, dict2 As Object) As Object
    Dim allKeys As Object: Set allKeys = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In dict1.Keys: allKeys(k) = True: Next k
    For Each k In dict2.Keys: allKeys(k) = True: Next k

    Dim n As Long: n = allKeys.Count
    Dim tk() As String: ReDim tk(1 To n)
    Dim idx As Long: idx = 1
    For Each k In allKeys.Keys
        tk(idx) = CStr(k): idx = idx + 1
    Next k

    Dim a As Long, b As Long, tmpS As String
    For a = 1 To n - 1
        For b = a + 1 To n
            If tk(b) < tk(a) Then
                tmpS = tk(a): tk(a) = tk(b): tk(b) = tmpS
            End If
        Next b
    Next a

    Dim pal(0 To 15) As Long
    pal(0)  = RGB(91, 155, 213)    ' blue
    pal(1)  = RGB(192, 80, 77)     ' red
    pal(2)  = RGB(155, 187, 89)    ' green
    pal(3)  = RGB(128, 100, 162)   ' purple
    pal(4)  = RGB(75, 172, 198)    ' teal
    pal(5)  = RGB(247, 150, 70)    ' orange
    pal(6)  = RGB(255, 217, 90)    ' gold
    pal(7)  = RGB(219, 219, 219)   ' light gray
    pal(8)  = RGB(146, 208, 80)    ' lime
    pal(9)  = RGB(0, 176, 240)     ' sky blue
    pal(10) = RGB(255, 105, 180)   ' pink
    pal(11) = RGB(191, 143, 0)     ' brown-gold
    pal(12) = RGB(255, 255, 102)   ' yellow
    pal(13) = RGB(0, 153, 153)     ' dark teal
    pal(14) = RGB(153, 102, 255)   ' violet
    pal(15) = RGB(200, 200, 200)   ' silver

    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    For idx = 1 To n
        result(tk(idx)) = pal((idx - 1) Mod 16)
    Next idx
    Set BuildTickerColorMap = result
End Function

' Sorts dict by value descending, then draws one pie chart named chartName
' at (leftPos, topPos), 400x320, styled via CombPerfPie.crtx (dark theme,
' Consolas fonts) exactly like the original single-pie implementation.
' colorMap forces each slice's fill by ticker so colors match across pies.
Private Sub RenderHoldingsPie(ws As Worksheet, dict As Object, chartName As String, _
    titleText As String, seriesName As String, leftPos As Double, topPos As Double, _
    colorMap As Object)

    If dict.Count = 0 Then Exit Sub

    ' Sort by value descending (insertion sort; small n)
    Dim n As Long: n = dict.Count
    Dim tickers() As String: ReDim tickers(1 To n)
    Dim vals() As Double:    ReDim vals(1 To n)
    Dim j As Long: j = 1
    Dim k As Variant
    For Each k In dict.Keys
        tickers(j) = CStr(k): vals(j) = CDbl(dict(k)): j = j + 1
    Next k
    Dim a As Long, b As Long, tmpD As Double, tmpS As String
    For a = 1 To n - 1
        For b = a + 1 To n
            If vals(b) > vals(a) Then
                tmpD = vals(a): vals(a) = vals(b): vals(b) = tmpD
                tmpS = tickers(a): tickers(a) = tickers(b): tickers(b) = tmpS
            End If
        Next b
    Next a

    Dim vArr() As Variant: ReDim vArr(1 To n)
    Dim lArr() As Variant: ReDim lArr(1 To n)
    For j = 1 To n: vArr(j) = vals(j): lArr(j) = tickers(j): Next j

    ' Remove old chart of the same name
    Dim co As ChartObject
    On Error Resume Next
    For Each co In ws.ChartObjects
        If co.Name = chartName Then co.Delete
    Next co
    On Error GoTo 0

    Dim pie As ChartObject
    Set pie = ws.ChartObjects.Add(leftPos, topPos, 400, 320)
    pie.Name = chartName

    Dim tplPath As String
    tplPath = Environ("USERPROFILE") & "\CombPerfPie.crtx"
    If Dir(tplPath) = "" Then tplPath = ""   ' fallback: style via VBA if file missing

    With pie.Chart
        .ChartType = 5   ' xlPie = 5

        ' Apply .crtx template: colours, fonts, dark theme all come from file
        On Error Resume Next
        If Dir(tplPath) <> "" Then .ApplyChartTemplate tplPath
        On Error GoTo 0

        ' Clear any placeholder series the template may have left,
        ' so our data becomes SeriesCollection(1) (pie only shows series 1)
        On Error Resume Next
        Dim delSC As Long
        For delSC = .SeriesCollection.Count To 1 Step -1
            .SeriesCollection(delSC).Delete
        Next delSC
        On Error GoTo 0

        ' Add our data as the sole series
        Dim sr As Series: Set sr = .SeriesCollection.NewSeries
        sr.Values  = vArr
        sr.XValues = lArr
        sr.Name    = seriesName

        ' Force each slice to its mapped color by ticker (not by slice index),
        ' so the same ticker renders identically in both pies.
        On Error Resume Next
        Dim pIdx As Long
        For pIdx = 1 To n
            If colorMap.Exists(tickers(pIdx)) Then
                sr.Points(pIdx).Interior.Color = colorMap(tickers(pIdx))
            End If
        Next pIdx
        On Error GoTo 0

        .HasTitle = True
        .ChartTitle.Text       = titleText
        .ChartTitle.Font.Bold  = True
        .ChartTitle.Font.Size  = 11
        .ChartTitle.Font.Name  = "Consolas"
        .ChartTitle.Font.Color = CLR_FG_WHITE

        .HasLegend = False   ' labels shown directly on each slice

        ' Data labels: ticker name + percentage on separate lines
        On Error Resume Next
        With .SeriesCollection(1)
            .ApplyDataLabels
            With .DataLabels
                .ShowCategoryName = True
                .ShowPercentage   = True
                .ShowValue        = False
                .Separator        = Chr(10)
                .NumberFormat     = "0%"
                .Font.Bold        = True
                .Font.Size        = 8
                .Font.Name        = "Consolas"
                .Font.Color       = CLR_FG_WHITE
            End With
        End With
        On Error GoTo 0
    End With
End Sub

Private Sub PieAddDict(dict As Object, key As String, val As Double)
    If dict.Exists(key) Then
        dict(key) = dict(key) + val
    Else
        dict(key) = val
    End If
End Sub

' Leveraged-ETF detector for RR4 holdings.
' TWSE/TPEx naming convention: a numeric ticker ending in "L" is always a
' 2x leveraged product (e.g. 00631L, 00675L = "positive 2x"); tickers ending
' in "R" are -1x inverse (same magnitude as unleveraged, so no adjustment).
' Extend the Select Case below if a non-Taiwan leveraged ETF (e.g. a US 3x
' product) ever shows up in RR4 and needs a manual multiplier.
Private Function GetLeverageMultiplier(ticker As String) As Double
    Dim tk As String: tk = UCase(Trim(ticker))
    GetLeverageMultiplier = 1
    If Len(tk) < 2 Then Exit Function
    If Right(tk, 1) = "L" And IsNumeric(Left(tk, Len(tk) - 1)) Then
        GetLeverageMultiplier = 2
        Exit Function
    End If
    Select Case tk
        ' Manual overrides for non-Taiwan leveraged ETFs (user's actual US holdings).
        Case "SOXL"                                   ' Direxion Semiconductor Bull 3x (sector index)
            GetLeverageMultiplier = 3
        Case "SNXX", "MRVU", "LABX", "SSO", "MUU"      ' single-stock / S&P500 2x products
            GetLeverageMultiplier = 2
    End Select
End Function


' =============================================================
' DIAGNOSTIC: scan RR4 Analysis sheet structure
' Run via Alt+F8 -> DiagnoseRR4
' Shows: sheet list, Analysis headers, all-col SUM scan, samples
' =============================================================
Public Sub DiagnoseRR4()
    Dim wb As Workbook
    Set wb = FindRR4WB()
    If wb Is Nothing Then
        MsgBox "RR4 workbook not found.  Open it first.", vbExclamation, "DiagnoseRR4"
        Exit Sub
    End If

    Dim out As String
    out = "=== RR4: " & wb.Name & " ===" & vbCrLf & vbCrLf

    ' -- List all sheet names --
    out = out & "[All Sheets]" & vbCrLf
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        out = out & "  " & ws.Name & vbCrLf
    Next ws
    out = out & vbCrLf

    ' -- Analysis sheet --
    Dim wsA As Worksheet
    On Error Resume Next
    Set wsA = wb.Sheets("Analysis")
    On Error GoTo 0

    If wsA Is Nothing Then
        out = out & "[Analysis] sheet NOT FOUND" & vbCrLf
        MsgBox out, vbExclamation, "DiagnoseRR4"
        Exit Sub
    End If

    ' Use UsedRange for true dimensions (avoids col=1 if row1 is a merged title)
    Dim ur As Range: Set ur = wsA.UsedRange
    Dim lastRow As Long: lastRow = ur.Row + ur.Rows.Count - 1
    Dim lastCol As Long: lastCol = ur.Column + ur.Columns.Count - 1
    out = out & "[Analysis]  usedRows=2-" & lastRow & "  usedCols=1-" & lastCol & vbCrLf & vbCrLf

    ' Dump every non-empty row: show all cols A..lastCol
    out = out & "--- All non-empty rows (full dump) ---" & vbCrLf
    Dim r As Long, c As Long
    For r = 1 To lastRow
        Dim rowStr As String: rowStr = ""
        For c = 1 To lastCol
            Dim sv As String: sv = Trim(CStr(wsA.Cells(r, c).Value))
            If sv <> "" Then
                rowStr = rowStr & "  " & ColLetter(c) & "=" & Left(sv, 16)
            End If
        Next c
        If rowStr <> "" Then
            out = out & "R" & r & ":" & rowStr & vbCrLf
        End If
    Next r

    ' Split output if too long
    If Len(out) > 1800 Then
        MsgBox Left(out, 1800), vbInformation, "DiagnoseRR4 (1/2)"
        MsgBox Mid(out, 1801), vbInformation, "DiagnoseRR4 (2/2)"
    Else
        MsgBox out, vbInformation, "DiagnoseRR4"
    End If
End Sub

Private Function ColLetter(n As Long) As String
    If n <= 26 Then
        ColLetter = Chr(64 + n)
    Else
        ColLetter = Chr(64 + (n - 1) \ 26) & Chr(65 + (n - 1) Mod 26)
    End If
End Function