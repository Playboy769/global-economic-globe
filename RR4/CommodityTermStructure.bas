Attribute VB_Name = "CommodityTermStructure"
Option Explicit

' ================================================================
'  ** DISABLED (2026-08-15) — no longer in active use, kept for
'  reference. Not wired to any button or event; only reachable by
'  manually running a macro from the VBE. **
' ================================================================

' ================================================================
'  COMMODITY TERM STRUCTURE — Bloomberg-Style Dashboard
'  v3.0  |  Dark Theme  |  Section 1: Curve Signal  |  Section 2: Spread Matrix
' ================================================================

' ── Sheet / layout constants ───────────────────────────────────
Private Const SH_MAIN    As String = "BBG_COMM"
Private Const SH_TICK    As String = "TickerMap"
Private Const SH_CFG     As String = "_Config"
Private Const MAX_RETRY  As Integer = 3

' ── Dashboard layout columns (BBG_COMM) ────────────────────────
Private Const COL_FIRST      As Long = 1
Private Const COL_LAST       As Long = 13  ' STRUCTURE column
Private Const COL_PREVCACHE  As Long = 14  ' hidden cache for prior-session M1
Private Const COL_UNUSED_GAP As Long = 12  ' visually reserved blank column

' ── Theme tokens (pure-black canvas, monospace) ────────────────
Private Const THEME_FONT      As String = "Consolas"
Private Const THEME_FONT_SIZE As Long = 10
Private Const C_BG            As Long = 0  ' RGB(0,0,0) - pure black canvas

' ================================================================
'  ENTRY POINT - refresh full dashboard
'  - TickerMap layout (V3): A=Instrument | B=Base | C=M1 | D=M2 | E=M3 | F=M6 | G=M12
'  - Computes % change vs previous close (cached in hidden column 14)
'  - Calls RefreshSpreadMatrix and writes timestamp
'  - Per-row error guard: a single failed fetch does not abort the run
' ================================================================
Sub RefreshDashboard()
    Dim wsM    As Worksheet
    Dim wsTick As Worksheet
    Dim wsCfg  As Worksheet
    Dim sec1   As Long, instCnt As Long, sec2Data As Long
    Dim i      As Long
    Const PREV_COL As Integer = 14   ' hidden cache for prior-session M1 close

    Set wsM = ThisWorkbook.Sheets(SH_MAIN)
    Set wsTick = ThisWorkbook.Sheets(SH_TICK)
    Set wsCfg = ThisWorkbook.Sheets(SH_CFG)

    sec1 = wsCfg.Range("B1").Value
    instCnt = wsCfg.Range("B2").Value
    sec2Data = wsCfg.Range("B3").Value

    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Commodity Term Structure..."
    wsM.Columns(PREV_COL).Hidden = True

    Dim okCount As Long, errCount As Long
    okCount = 0: errCount = 0

    For i = 1 To instCnt
        Dim dataRow As Long
        dataRow = sec1 + i - 1

        Application.StatusBar = "Fetching [" & i & "/" & instCnt & "]  " & _
                                  wsM.Cells(dataRow, 2).Value

        ' Per-row error guard so a single failed fetch cannot abort the whole run.
        On Error GoTo RowFail

        ' V3 TickerMap layout: C..G = M1, M2, M3, M6, M12.
        Dim tM1  As String, tM2  As String, tM3  As String
        Dim tM6  As String, tM12 As String
        tM1 = Trim(wsTick.Cells(i + 1, 3).Value)
        tM2 = Trim(wsTick.Cells(i + 1, 4).Value)
        tM3 = Trim(wsTick.Cells(i + 1, 5).Value)
        tM6 = Trim(wsTick.Cells(i + 1, 6).Value)
        tM12 = Trim(wsTick.Cells(i + 1, 7).Value)

        Dim baseSym As String
        baseSym = Trim(wsTick.Cells(i + 1, 2).Value)

        Dim pM1 As Double, pM2 As Double, pM3 As Double
        Dim pM6 As Double, pM12 As Double
        pM1 = FetchPriceWithFallback(tM1, baseSym)
        pM2 = FetchPriceWithFallback(tM2, baseSym)
        pM3 = FetchPriceWithFallback(tM3, baseSym)
        pM6 = FetchPriceWithFallback(tM6, baseSym)
        pM12 = FetchPriceWithFallback(tM12, baseSym)

        If pM1 > 0 Then
            ' % change vs cached prior M1 close.
            Dim prevClose As Double: prevClose = 0
            If IsNumeric(wsM.Cells(dataRow, PREV_COL).Value) Then
                prevClose = CDbl(wsM.Cells(dataRow, PREV_COL).Value)
            End If
            With wsM.Cells(dataRow, 3)
                .Font.Name = THEME_FONT
                .Font.Size = THEME_FONT_SIZE
                .Interior.Color = C_BG
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                If prevClose > 0 Then
                    Dim chg As Double: chg = (pM1 - prevClose) / prevClose
                    .Value = chg
                    .NumberFormat = "+0.00%;-0.00%;0.00%"
                    .Font.Bold = True
                    If chg > 0 Then
                        .Font.Color = RGB(0, 255, 136)
                    ElseIf chg < 0 Then
                        .Font.Color = RGB(255, 51, 51)
                    Else
                        .Font.Color = RGB(255, 255, 0)
                    End If
                Else
                    .Value = "new"
                    .Font.Color = RGB(136, 136, 136)
                End If
            End With
            wsM.Cells(dataRow, PREV_COL).Value = pM1

            ' Prices
            WritePrice wsM.Cells(dataRow, 4), pM1
            WritePrice wsM.Cells(dataRow, 5), pM2
            WritePrice wsM.Cells(dataRow, 6), pM3
            WritePrice wsM.Cells(dataRow, 7), pM6
            WritePrice wsM.Cells(dataRow, 8), pM12

            ' Spreads (guard against zero far-month price -> blank, not bogus spread)
            If pM2 > 0 Then WriteSpread wsM.Cells(dataRow, 9), pM1 - pM2 _
                       Else wsM.Cells(dataRow, 9).Value = "-"
            If pM6 > 0 Then WriteSpread wsM.Cells(dataRow, 10), pM1 - pM6 _
                       Else wsM.Cells(dataRow, 10).Value = "-"
            If pM12 > 0 Then WriteSpread wsM.Cells(dataRow, 11), pM1 - pM12 _
                        Else wsM.Cells(dataRow, 11).Value = "-"

            ' Structure: prefer M12, fall back to M6 then M2 when far months unavailable.
            Dim farPrice As Double
            If pM12 > 0 Then
                farPrice = pM12
            ElseIf pM6 > 0 Then
                farPrice = pM6
            Else
                farPrice = pM2
            End If
            If farPrice > 0 Then ApplyStructure wsM.Cells(dataRow, 13), pM1, farPrice

            okCount = okCount + 1
        Else
            MarkError wsM, dataRow
            errCount = errCount + 1
        End If

        On Error GoTo 0
        GoTo RowDone
RowFail:
        MarkError wsM, dataRow
        errCount = errCount + 1
        On Error GoTo 0
RowDone:
    Next i

    ' Wipe stray rows between the last Section-1 instrument and the
    ' Section-2 title band. Stop at sec2Data - 3 so we never touch:
    '   sec2Data - 2  (Section 2 title row, merged A:M)
    '   sec2Data - 1  (Section 2 header row)
    ' Unmerge each row first so the wipe cannot trip "Cannot change
    ' part of a merged cell" (Excel error 1004).
    Dim wipeRow As Long, cc As Integer
    For wipeRow = sec1 + instCnt To sec2Data - 3
        On Error Resume Next
        wsM.Range(wsM.Cells(wipeRow, 1), wsM.Cells(wipeRow, COL_LAST)).UnMerge
        On Error GoTo 0
        For cc = 1 To 13
            With wsM.Cells(wipeRow, cc)
                .ClearContents
                .Interior.Color = C_BG
                .Font.Color = C_BG
                .Font.Name = THEME_FONT
                .Font.Size = THEME_FONT_SIZE
            End With
        Next cc
    Next wipeRow

    ' Section 2: Calendar Spread Matrix (reads selected row from _Config!B5).
    Call RefreshSpreadMatrix(wsM, wsTick, wsCfg, sec1, sec2Data)

    ' Timestamp: spread matrix occupies 6 rows -> first free row is sec2Data + 6.
    Dim tsRow As Long
    tsRow = sec2Data + 6
    With wsM.Cells(tsRow, 2)
        .Value = "Last updated: " & Format(Now(), "yyyy/mm/dd  hh:mm:ss")
        .Font.Name = THEME_FONT
        .Font.Size = 8
        .Font.Color = RGB(100, 100, 100)
        .Interior.Color = C_BG
        .HorizontalAlignment = xlLeft
    End With

    Application.ScreenUpdating = True
    Application.StatusBar = "Dashboard refreshed - " & okCount & " OK / " & errCount & " errors"
End Sub

' Backward-compat alias for any toolbar button still pointing at V2.
Sub RefreshDashboardV2()
    Call RefreshDashboard
End Sub

Private Sub RefreshSpreadMatrix(wsM As Worksheet, wsTick As Worksheet, _
                                 wsCfg As Worksheet, sec1 As Long, sec2Data As Long)
    ' Resolve which Section-1 instrument feeds the spread matrix.
    ' _Config!B5 stores a 1-based offset within Section 1 (1 = first instrument).
    ' Defaults to 1 if missing / out of range / non-numeric.
    Dim instCnt As Long
    instCnt = CLng(wsCfg.Range("B2").Value)
    Dim selOffset As Long: selOffset = 1
    If IsNumeric(wsCfg.Range("B5").Value) Then
        selOffset = CLng(wsCfg.Range("B5").Value)
        If selOffset < 1 Or selOffset > instCnt Then selOffset = 1
    End If
    Dim selRow As Long: selRow = sec1 + selOffset - 1

    ' Robustness: if the user-selected row has no usable M1 price, scan
    ' forward from sec1 for the first row that does. Avoids the matrix
    ' silently displaying all-zero spreads when _Config is stale.
    If Not RowHasValidPrices(wsM, selRow) Then
        Dim probe As Long
        For probe = sec1 To sec1 + instCnt - 1
            If RowHasValidPrices(wsM, probe) Then
                selRow = probe
                Exit For
            End If
        Next probe
    End If

    ' Echo the selected instrument name into the section header so the user
    ' can see which curve is being decomposed.
    On Error Resume Next
    With wsM.Cells(sec2Data - 1, 7)
        .Value = "Selected: " & wsM.Cells(selRow, 2).Value
        .Font.Name = THEME_FONT
        .Font.Size = THEME_FONT_SIZE - 1
        .Font.Color = RGB(180, 180, 180)
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With
    On Error GoTo 0

    ' Read M1..M12 prices (BBG_COMM columns 4..8).
    Dim prices(1 To 5) As Double
    Dim c As Integer
    For c = 1 To 5
        Dim v As Variant: v = wsM.Cells(selRow, 3 + c).Value
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then prices(c) = CDbl(v) Else prices(c) = 0
        Else
            prices(c) = 0
        End If
    Next c

    ' Populate the TICKERS column (col 3 in the matrix) from the
    ' TickerMap row that corresponds to selRow.
    Dim tickRow As Long: tickRow = (selRow - sec1) + 2   ' +2: header row at 1, data starts at 2
    Dim labels As Variant
    labels = Array("M1-M2", "M1-M3", "M1-M6", "M1-M12", "M2-M3", "M3-M6")
    Dim tickCols As Variant
    ' Column indexes in TickerMap (V3 layout): C=M1, D=M2, E=M3, F=M6, G=M12
    tickCols = Array(Array(3, 4), Array(3, 5), Array(3, 6), Array(3, 7), Array(4, 5), Array(5, 6))
    Dim jj As Integer
    For jj = 0 To 5
        Dim ta As String, tb As String
        On Error Resume Next
        ta = CStr(wsTick.Cells(tickRow, tickCols(jj)(0)).Value)
        tb = CStr(wsTick.Cells(tickRow, tickCols(jj)(1)).Value)
        On Error GoTo 0
        With wsM.Cells(sec2Data + jj, 3)
            .Value = ta & " / " & tb
            .Font.Name = THEME_FONT
            .Font.Size = 8
            .Font.Color = RGB(140, 140, 140)
            .Interior.Color = C_BG
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next jj

    Dim spreads(1 To 6) As Double
    spreads(1) = prices(1) - prices(2)
    spreads(2) = prices(1) - prices(3)
    spreads(3) = prices(1) - prices(4)
    spreads(4) = prices(1) - prices(5)
    spreads(5) = prices(2) - prices(3)
    spreads(6) = prices(3) - prices(4)

    Dim j As Integer
    For j = 1 To 6
        Dim r As Long: r = sec2Data + j - 1

        ' �w�w ���Ѱ��ӦC�Ҧ��X�֡A�קK 1004 �w�w�w�w�w�w�w�w�w�w�w�w�w�w
        On Error Resume Next
        wsM.Rows(r).UnMerge
        On Error GoTo 0

        Dim spd As Double: spd = spreads(j)

        ' Value (col 4)
        WriteSpread wsM.Cells(r, 4), spd

        ' Direction (col 5)
        With wsM.Cells(r, 5)
            ApplyMatrixCellTheme wsM.Cells(r, 5)
            .Font.Bold = True
            If spd > 0.005 Then
                .Value = "[+] BKWD"
                .Font.Color = RGB(0, 255, 136)
            ElseIf spd < -0.005 Then
                .Value = "[-] CNTNG"
                .Font.Color = RGB(255, 51, 51)
            Else
                .Value = "[=] FLAT"
                .Font.Color = RGB(255, 255, 0)
            End If
        End With

        ' Strength (col 6)
        Dim pctSpd As Double
        If prices(1) > 0 Then pctSpd = Abs(spd) / prices(1) Else pctSpd = 0
        With wsM.Cells(r, 6)
            ApplyMatrixCellTheme wsM.Cells(r, 6)
            .Font.Bold = True
            If pctSpd > 0.05 Then
                .Value = "STRONG": .Font.Color = RGB(0, 255, 136)
            ElseIf pctSpd > 0.01 Then
                .Value = "MODERATE": .Font.Color = RGB(255, 192, 0)
            Else
                .Value = "WEAK": .Font.Color = RGB(136, 136, 136)
            End If
        End With

        ' Interpretation (col 7)
        With wsM.Cells(r, 7)
            ApplyMatrixCellTheme wsM.Cells(r, 7)
            .Font.Bold = True
            If spd > 0.005 Then
                .Value = "TIGHT SUPPLY / STRONG SPOT DEMAND"
                .Font.Color = RGB(0, 255, 136)
            ElseIf spd < -0.005 Then
                .Value = "STORAGE SAT / FULL CARRY"
                .Font.Color = RGB(255, 51, 51)
            Else
                .Value = "MARKET IN EQUILIBRIUM"
                .Font.Color = RGB(255, 255, 0)
            End If
        End With

        ' Roll Yield % (col 8)
        With wsM.Cells(r, 8)
            ApplyMatrixCellTheme wsM.Cells(r, 8)
            .NumberFormat = "0.00%"
            .Font.Bold = True
            If prices(1) > 0 Then
                .Value = spd / prices(1)
                If spd >= 0 Then
                    .Font.Color = RGB(0, 255, 136)
                Else
                    .Font.Color = RGB(255, 51, 51)
                End If
            Else
                .Value = "-"
                .Font.Color = RGB(136, 136, 136)
            End If
        End With

        ' Signal Flag (col 13)
        With wsM.Cells(r, 13)
            ApplyMatrixCellTheme wsM.Cells(r, 13)
            .Font.Bold = True
            If Abs(spd) > 5 Then
                .Value = ">> WATCH": .Font.Color = RGB(255, 51, 51)
            ElseIf Abs(spd) > 1 Then
                .Value = "!! SHARP MOVE": .Font.Color = RGB(255, 192, 0)
            Else
                .Value = "OK STABLE": .Font.Color = RGB(136, 136, 136)
            End If
        End With

        ' Reserved gap col 12 stays clean.
        With wsM.Cells(r, COL_UNUSED_GAP)
            .ClearContents
            .Interior.Color = C_BG
        End With
    Next j
End Sub

' Shared per-cell theme for Section-2 matrix cells (Consolas / centered / black bg).
Private Sub ApplyMatrixCellTheme(cell As Range)
    cell.Font.Name = THEME_FONT
    cell.Font.Size = THEME_FONT_SIZE
    cell.Interior.Color = C_BG
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
End Sub

' ================================================================
'  ApplyDarkTheme
'  One-shot UI sweep for BBG_COMM:
'    - Paint the entire used range pure black
'    - Set every cell font to Consolas
'    - Re-render Section 1 / Section 2 title bands (merged across A:M)
'    - Re-render legend band at top (rows 2-3)
'    - Center the instrument column (B) left, headers center
'    - Hide the prev-close cache column (14)
'  Run this once after any layout change, or simply re-run; it is idempotent.
' ================================================================
Sub ApplyDarkTheme()
    Dim wsM   As Worksheet: Set wsM = ThisWorkbook.Sheets(SH_MAIN)
    Dim wsCfg As Worksheet: Set wsCfg = ThisWorkbook.Sheets(SH_CFG)

    Dim sec1     As Long: sec1 = CLng(wsCfg.Range("B1").Value)
    Dim instCnt  As Long: instCnt = CLng(wsCfg.Range("B2").Value)
    Dim sec2Data As Long: sec2Data = CLng(wsCfg.Range("B3").Value)

    Application.ScreenUpdating = False

    ' Kill default Excel gridlines on this sheet so the pure-black canvas
    ' has no light-grey artefact lines bleeding through.
    On Error Resume Next
    wsM.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    On Error GoTo 0

    ' --- 1. Paint everything pure black + Consolas across the dashboard band.
    Dim lastRow As Long: lastRow = sec2Data + 8
    Dim canvas As Range
    Set canvas = wsM.Range(wsM.Cells(1, 1), wsM.Cells(lastRow, COL_LAST))
    With canvas
        .Interior.Color = C_BG
        .Font.Name = THEME_FONT
        .Font.Size = THEME_FONT_SIZE
        .Font.Color = RGB(220, 220, 220)
        .VerticalAlignment = xlCenter
        ' Strip any leftover borders so frame color is fully managed by us.
        .Borders.LineStyle = xlNone
    End With

    ' Whole-sheet sweep so neighbouring cells (units, helper text) also get Consolas.
    With wsM.Cells
        .Font.Name = THEME_FONT
    End With

    ' --- 2. Legend band (rows 2-3). Merge B:M (column A stays as gutter).
    UnmergeBand wsM, 2, 3
    With wsM.Range(wsM.Cells(2, 2), wsM.Cells(2, COL_LAST))
        .Merge
        .Value = "BACKWARDATION  ->  Supply tightness. No storage incentive. Physical scarcity drives spot premium over forward."
        .Font.Name = THEME_FONT
        .Font.Color = RGB(0, 255, 136)
        .Font.Italic = True
        .Font.Size = THEME_FONT_SIZE
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders.LineStyle = xlNone
    End With
    With wsM.Range(wsM.Cells(3, 2), wsM.Cells(3, COL_LAST))
        .Merge
        .Value = "CONTANGO       ->  Oversupply. Storage economical - buy spot cheap, sell forward at premium. Physical demand weak."
        .Font.Name = THEME_FONT
        .Font.Color = RGB(255, 51, 51)
        .Font.Italic = True
        .Font.Size = THEME_FONT_SIZE
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders.LineStyle = xlNone
    End With

    ' --- 3. Section 1 title band (row sec1 - 2) and header row (sec1 - 1).
    '        Title merges B:M (column A stays as gutter).
    Dim sec1Title As Long: sec1Title = sec1 - 2
    Dim sec1Head  As Long: sec1Head = sec1 - 1
    If sec1Title >= 4 Then
        UnmergeBand wsM, sec1Title, sec1Title
        With wsM.Range(wsM.Cells(sec1Title, 2), wsM.Cells(sec1Title, COL_LAST))
            .Merge
            .Value = "< SECTION 1 >   CURVE STRUCTURE SIGNAL   -   FRONT MONTH vs DEFERRED MONTHS"
            .Font.Name = THEME_FONT
            .Font.Color = RGB(220, 220, 220)
            .Font.Bold = True
            .Font.Size = THEME_FONT_SIZE + 1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(20, 20, 20)
            .Borders.LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(60, 60, 60)
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    End If
    PaintHeaderRow wsM, sec1Head, _
        Array("", "INSTRUMENT", "% Change", "M1", "M2", "M3", "M6", "M12", _
              "M1-M2 SPRD", "M1-M6 SPRD", "M1-M12 SPRD", "", "STRUCTURE")

    ' --- 4. Section 1 instrument column (B) left-align bold, dim units.
    Dim i As Long
    For i = 0 To instCnt - 1
        With wsM.Cells(sec1 + i, 2)
            .Font.Name = THEME_FONT
            .Font.Size = THEME_FONT_SIZE
            .Font.Bold = True
            .Font.Color = RGB(220, 220, 220)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
    Next i

    ' --- 5. Section 2 title + header. Title merges B:M (column A stays gutter).
    Dim sec2Title As Long: sec2Title = sec2Data - 2
    Dim sec2Head  As Long: sec2Head = sec2Data - 1
    If sec2Title >= sec1 + instCnt Then
        UnmergeBand wsM, sec2Title, sec2Title
        With wsM.Range(wsM.Cells(sec2Title, 2), wsM.Cells(sec2Title, COL_LAST))
            .Merge
            .Value = "< SECTION 2 >   CALENDAR SPREAD MATRIX   -   SELECT INSTRUMENT VIA _Config!B5"
            .Font.Name = THEME_FONT
            .Font.Color = RGB(220, 220, 220)
            .Font.Bold = True
            .Font.Size = THEME_FONT_SIZE + 1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(20, 20, 20)
            .Borders.LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(60, 60, 60)
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    End If
    PaintHeaderRow wsM, sec2Head, _
        Array("", "SPREAD PERIOD", "TICKERS", "VALUE", "DIRECTION", "STRENGTH", _
              "INTERPRETATION", "ROLL YIELD %", "PRIOR WEEK", "WoW CHANGE", "MOMENTUM", "", "SIGNAL FLAG")

    ' --- 6. Section 2 first column (SPREAD PERIOD labels).
    Dim sLabels As Variant
    sLabels = Array("M1 - M2  [ 1mo ]", "M1 - M3  [ 2mo ]", _
                    "M1 - M6  [ 5mo ]", "M1 - M12 [11mo ]", _
                    "M2 - M3  [ roll ]", "M3 - M6  [ roll ]")
    Dim k As Long
    For k = 0 To 5
        With wsM.Cells(sec2Data + k, 2)
            .Value = sLabels(k)
            .Font.Name = THEME_FONT
            .Font.Size = THEME_FONT_SIZE
            .Font.Bold = True
            .Font.Color = RGB(220, 220, 220)
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        End With
    Next k

    ' --- 7. Unified column widths sized for the LONGEST content per column.
    '   A : 3   - left gutter (visible margin so row-2/3/5 merges read as "from col B")
    '   B : 26  - INSTRUMENT names  / SPREAD PERIOD labels
    '   C : 22  - TICKERS pair "CLM26.NYM / CLM26.NYM" (Section 2) / % Change (Section 1)
    '   D : 12  - M1  / VALUE
    '   E : 13  - M2  / DIRECTION  ("[+] BKWD")
    '   F : 11  - M3  / STRENGTH   ("MODERATE")
    '   G : 36  - M6  / INTERPRETATION ("TIGHT SUPPLY / STRONG SPOT DEMAND" = 33)
    '   H : 13  - M12 / ROLL YIELD %
    '   I : 13  - M1-M2 SPRD / PRIOR WEEK
    '   J : 13  - M1-M6 SPRD / WoW CHANGE
    '   K : 13  - M1-M12 SPRD / MOMENTUM
    '   L : 3   - reserved gap
    '   M : 22  - STRUCTURE ("BACKWARDATION") / SIGNAL FLAG ("!! SHARP MOVE")
    wsM.Columns(1).ColumnWidth = 3
    wsM.Columns(2).ColumnWidth = 26
    wsM.Columns(3).ColumnWidth = 22
    wsM.Columns(4).ColumnWidth = 12
    wsM.Columns(5).ColumnWidth = 13
    wsM.Columns(6).ColumnWidth = 11
    wsM.Columns(7).ColumnWidth = 36
    wsM.Columns(8).ColumnWidth = 13
    wsM.Columns(9).ColumnWidth = 13
    wsM.Columns(10).ColumnWidth = 13
    wsM.Columns(11).ColumnWidth = 13
    wsM.Columns(COL_UNUSED_GAP).ColumnWidth = 3
    wsM.Columns(COL_LAST).ColumnWidth = 22
    wsM.Columns(COL_PREVCACHE).Hidden = True

    ' --- 7b. Uniform row heights. Rows 21-23 came out short because
    ' SyncLayoutFromTickerMap appended new instrument rows without resetting
    ' the row height. Normalise every data row and the title bands.
    Dim rh As Long
    For rh = sec1 To sec1 + instCnt - 1
        wsM.Rows(rh).RowHeight = 20
    Next rh
    For rh = sec2Data To sec2Data + 5
        wsM.Rows(rh).RowHeight = 20
    Next rh
    wsM.Rows(2).RowHeight = 18
    wsM.Rows(3).RowHeight = 18
    If sec1Title >= 4 Then wsM.Rows(sec1Title).RowHeight = 24
    If sec2Title >= sec1 + instCnt Then wsM.Rows(sec2Title).RowHeight = 24
    wsM.Rows(sec1Head).RowHeight = 22
    wsM.Rows(sec2Head).RowHeight = 22

    ' --- 8. Freeze panes at first data row of Section 1 so headers stay visible.
    On Error Resume Next
    wsM.Activate
    wsM.Cells(sec1, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' --- 9. Sheet tab color: Bloomberg-style amber so BBG_COMM stands out
    ' in the workbook tab bar.
    On Error Resume Next
    wsM.Tab.Color = RGB(230, 170, 40)
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "Dark theme applied. Run RefreshDashboard to repopulate prices.", vbInformation
End Sub

' Helper - unmerge a horizontal band before re-merging (idempotent).
' Sweeps a wide range so legacy merges that include column A or extend past
' COL_LAST are also broken before the new merge is applied.
Private Sub UnmergeBand(wsM As Worksheet, rTop As Long, rBot As Long)
    On Error Resume Next
    wsM.Range(wsM.Cells(rTop, 1), wsM.Cells(rBot, COL_LAST + 4)).UnMerge
    On Error GoTo 0
End Sub

' Helper - paint a header row from an array of labels (COL_LAST cells wide).
Private Sub PaintHeaderRow(wsM As Worksheet, rowIdx As Long, labels As Variant)
    Dim c As Long
    For c = 1 To COL_LAST
        With wsM.Cells(rowIdx, c)
            .Value = labels(c - 1)
            .Font.Name = THEME_FONT
            .Font.Size = THEME_FONT_SIZE
            .Font.Bold = True
            .Font.Color = RGB(180, 180, 180)
            .Interior.Color = RGB(25, 25, 25)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(80, 80, 80)
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    Next c
End Sub
' ================================================================
'  Price fetcher with dual-endpoint fallback + retry
' ================================================================
Private Function FetchPrice(symbol As String) As Double
    On Error GoTo Bail
    If symbol = "" Then FetchPrice = 0: Exit Function
    Dim attempt As Integer, result As Double
    For attempt = 1 To MAX_RETRY
        result = FetchYahoo(symbol)
        If result > 0 Then FetchPrice = result: Exit Function
        Call PauseMS(600 * attempt)
    Next attempt
    FetchPrice = 0
    Exit Function
Bail:
    FetchPrice = 0
End Function

' Tries the per-month ticker first, then falls back to the front-month
' continuous contract for this base if the monthly symbol returns nothing.
' Yahoo's coverage of CME monthly tickers (e.g. CLM26.NYM, ESH26.CME) is
' inconsistent, so this fallback keeps the dashboard populated rather than
' showing ERR for every far-month contract.
'
' Equity index futures (ES/NQ/YM/RTY/VX) are short-circuited: Yahoo only
' carries the continuous symbol for these, so we skip the monthly lookup
' entirely to avoid wasting 3 retries per cell on a guaranteed 404.
Private Function FetchPriceWithFallback(symbol As String, baseSym As String) As Double
    If IsIndexFutureBase(baseSym) Then
        FetchPriceWithFallback = FetchPrice(FrontMonthYahooSymbol(baseSym))
        Exit Function
    End If

    Dim p As Double
    p = FetchPrice(symbol)
    If p > 0 Then FetchPriceWithFallback = p: Exit Function

    If LenB(baseSym) > 0 Then
        Dim alt As String: alt = FrontMonthYahooSymbol(baseSym)
        ' Avoid an infinite loop if caller already passed the same form.
        If StrComp(alt, symbol, vbTextCompare) <> 0 Then
            p = FetchPrice(alt)
            If p > 0 Then FetchPriceWithFallback = p: Exit Function
        End If
    End If
    FetchPriceWithFallback = 0
End Function

' Whitelist of equity index futures whose individual-month contracts are
' NOT carried by Yahoo Finance. Treat as front-month-only.
Private Function IsIndexFutureBase(baseSym As String) As Boolean
    Select Case UCase(Trim(baseSym))
        Case "ES", "NQ", "YM", "RTY", "VX"
            IsIndexFutureBase = True
        Case Else
            IsIndexFutureBase = False
    End Select
End Function

' Resolves the Yahoo symbol that actually serves data for this base's
' front-month / spot quote. Most futures use <base>=F, but a handful of
' instruments need a special route:
'   VX  -> ^VIX  (Yahoo has no VX=F; ^VIX is the VIX spot index, which is
'                 the closest publicly available proxy for VX1 front-month)
Private Function FrontMonthYahooSymbol(baseSym As String) As String
    Select Case UCase(Trim(baseSym))
        Case "VX"
            FrontMonthYahooSymbol = "^VIX"
        Case Else
            FrontMonthYahooSymbol = baseSym & "=F"
    End Select
End Function

' True when col 4 (M1 price) is a positive number on a Section-1 row.
' Used by RefreshSpreadMatrix to skip rows that failed to fetch.
Private Function RowHasValidPrices(wsM As Worksheet, r As Long) As Boolean
    On Error GoTo Bail
    Dim v As Variant: v = wsM.Cells(r, 4).Value
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then RowHasValidPrices = True: Exit Function
    End If
Bail:
    RowHasValidPrices = False
End Function

Sub FixAndRebuild()
    ' �w�w �B�J1�G���� _Config �ƭ� �w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w
    Dim wsCfg As Worksheet
    On Error Resume Next
    Set wsCfg = ThisWorkbook.Sheets("_Config")
    On Error GoTo 0
    If wsCfg Is Nothing Then
        Set wsCfg = ThisWorkbook.Sheets.Add
        wsCfg.Name = "_Config"
        wsCfg.Visible = xlSheetHidden
    End If
    
    ' �۰ʰ��� BBG_COMM �u�@������ڵ��c
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets("BBG_COMM")
    
    ' ���y�� Section1 �_�l�C�]��uINSTRUMENT�v���D�^
    Dim sec1 As Long, i As Long
    For i = 1 To 50
        If InStr(wsM.Cells(i, 2).Value, "INSTRUMENT") > 0 Then
            sec1 = i + 1   ' ��Ʊq���D�U�@�C�}�l
            Exit For
        End If
    Next i
    If sec1 = 0 Then sec1 = 7   ' fallback �w�]��
    
    ' ���y�� Section2 �_�l�C�]��uSPREAD PERIOD�v���D�^
    Dim sec2Hdr As Long, sec2Data As Long
    For i = sec1 To 100
        If InStr(wsM.Cells(i, 2).Value, "SPREAD PERIOD") > 0 Then
            sec2Hdr = i
            sec2Data = i + 1
            Exit For
        End If
    Next i
    If sec2Data = 0 Then sec2Data = sec1 + 20   ' fallback
    
    ' �p��ӫ~�ƶq�]sec1 �� sec2Hdr ���������e���C�^
    Dim instCnt As Long
    For i = sec1 To sec2Hdr - 1
        If Trim(wsM.Cells(i, 2).Value) <> "" Then instCnt = instCnt + 1
    Next i
    If instCnt = 0 Then instCnt = 14   ' fallback
    
    ' �g�J _Config
    wsCfg.Range("A1") = "SEC1_START": wsCfg.Range("B1") = sec1
    wsCfg.Range("A2") = "INST_COUNT": wsCfg.Range("B2") = instCnt
    wsCfg.Range("A3") = "SEC2_DATA_START": wsCfg.Range("B3") = sec2Data
    ' Spread-matrix instrument selector (1-based offset within Section 1).
    If Not IsNumeric(wsCfg.Range("B5").Value) Then
        wsCfg.Range("A5") = "SpreadMatrix_Offset": wsCfg.Range("B5") = 1
    End If
    
    MsgBox "? _Config �w�״_�I" & vbCr & vbCr & _
           "sec1      = " & sec1 & vbCr & _
           "instCnt   = " & instCnt & vbCr & _
           "sec2Data  = " & sec2Data & vbCr & vbCr & _
           "�{�b�i�H���� RefreshDashboard", vbInformation
End Sub
Private Function FetchYahoo(symbol As String) As Double
    ' Full error trap: any HTTP / parse failure returns 0 instead of bubbling
    ' up to the caller. A fresh MSXML2.XMLHTTP is created per endpoint so a
    ' stuck state on attempt #1 cannot poison attempt #2.
    On Error GoTo Bail

    ' URL-encode the few characters that appear in Yahoo symbols but break
    ' raw URL parsing in MSXML. Most important: ^VIX -> %5EVIX.
    Dim safeSym As String: safeSym = symbol
    safeSym = Replace(safeSym, "^", "%5E")

    Dim eps(1) As String
    eps(0) = "https://query1.finance.yahoo.com/v8/finance/chart/" & safeSym & "?interval=1d&range=1d"
    eps(1) = "https://query2.finance.yahoo.com/v8/finance/chart/" & safeSym & "?interval=1d&range=1d"

    Dim ep As Integer
    For ep = 0 To 1
        Dim http As Object
        Set http = Nothing
        Set http = CreateObject("MSXML2.XMLHTTP")

        Dim sent As Boolean: sent = False
        On Error Resume Next
        http.Open "GET", eps(ep), False
        http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0)"
        http.send
        sent = (Err.Number = 0)
        On Error GoTo Bail

        If sent Then
            Dim status As Long: status = 0
            On Error Resume Next
            status = http.status
            On Error GoTo Bail

            If status = 200 Then
                Dim response As String: response = ""
                On Error Resume Next
                response = http.responseText
                On Error GoTo Bail

                Dim parsed As Double
                parsed = ParseYahooPrice(response)
                If parsed > 0 Then FetchYahoo = parsed: Exit Function
            End If
        End If
        Set http = Nothing
    Next ep

    FetchYahoo = 0
    Exit Function
Bail:
    FetchYahoo = 0
End Function

' Pulls regularMarketPrice from the Yahoo chart JSON, with a fallback to
' the last numeric value in the close[] array.
Private Function ParseYahooPrice(response As String) As Double
    On Error GoTo Bail
    If LenB(response) = 0 Then ParseYahooPrice = 0: Exit Function

    Dim pos As Long
    pos = InStr(response, """regularMarketPrice"":")
    If pos > 0 Then
        Dim tmp As String
        tmp = Mid(response, pos + 22, 20)
        tmp = Split(tmp, ",")(0)
        tmp = Split(tmp, "}")(0)
        tmp = Trim(tmp)
        If IsNumeric(tmp) Then
            If CDbl(tmp) > 0 Then ParseYahooPrice = CDbl(tmp): Exit Function
        End If
    End If

    pos = InStr(response, """close"":[")
    If pos > 0 Then
        Dim arr As String
        arr = Mid(response, pos + 9, 300)
        arr = Split(arr, "]")(0)
        Dim parts() As String: parts = Split(arr, ",")
        ' Walk from the tail backwards, skipping nulls that Yahoo sometimes
        ' emits for the most recent (still-open) bar.
        Dim k As Long
        For k = UBound(parts) To LBound(parts) Step -1
            Dim cand As String: cand = Trim(parts(k))
            If IsNumeric(cand) Then
                If CDbl(cand) > 0 Then ParseYahooPrice = CDbl(cand): Exit Function
            End If
        Next k
    End If

    ParseYahooPrice = 0
    Exit Function
Bail:
    ParseYahooPrice = 0
End Function

Sub FixTickerMap()
    Dim wsTick As Worksheet
    On Error Resume Next
    Set wsTick = ThisWorkbook.Sheets("TickerMap")
    On Error GoTo 0
    If wsTick Is Nothing Then
        Set wsTick = ThisWorkbook.Sheets.Add
        wsTick.Name = "TickerMap"
        wsTick.Visible = xlSheetHidden
    End If
    
    wsTick.Cells.Clear
    
    ' ���D
    Dim h As Variant
    Dim headers As Variant
    headers = Array("Instrument", "M1", "M2", "M3", "M6", "M12")
    Dim ci As Integer
    For ci = 0 To 5
        wsTick.Cells(1, ci + 1).Value = headers(ci)
    Next ci
    
    ' �w�w Yahoo Finance �s��X���u�� =F�]���^�A
    '    ����� +1/+2/+5/+11 �몺����N�X
    '    �榡�GCLK25.NYM, CLM25.NYM... ��ڤW Yahoo ���䴩
    '    �̥i�a���k�G�����ΦP�@�� =F �N�X�A�� VBA ������
    '    ����ȥ� Quandl/CME �榡�A�Χ�� M2 spread �Ϊ�����
    
    Dim data As Variant
    data = Array( _
        Array("WTI CRUDE OIL", "CL=F", "CL=F", "CL=F", "CL=F", "CL=F"), _
        Array("BRENT CRUDE", "BZ=F", "BZ=F", "BZ=F", "BZ=F", "BZ=F"), _
        Array("RBOB GASOLINE", "RB=F", "RB=F", "RB=F", "RB=F", "RB=F"), _
        Array("ULSD DIESEL", "HO=F", "HO=F", "HO=F", "HO=F", "HO=F"), _
        Array("NATURAL GAS", "NG=F", "NG=F", "NG=F", "NG=F", "NG=F"), _
        Array("GOLD", "GC=F", "GC=F", "GC=F", "GC=F", "GC=F"), _
        Array("COPPER", "HG=F", "HG=F", "HG=F", "HG=F", "HG=F"), _
        Array("SILVER", "SI=F", "SI=F", "SI=F", "SI=F", "SI=F"), _
        Array("CORN", "ZC=F", "ZC=F", "ZC=F", "ZC=F", "ZC=F"), _
        Array("SOYBEANS", "ZS=F", "ZS=F", "ZS=F", "ZS=F", "ZS=F"), _
        Array("WHEAT", "ZW=F", "ZW=F", "ZW=F", "ZW=F", "ZW=F"), _
        Array("ALUMINUM", "ALI=F", "ALI=F", "ALI=F", "ALI=F", "ALI=F"), _
        Array("LUMBER", "LBS=F", "LBS=F", "LBS=F", "LBS=F", "LBS=F"), _
        Array("STEEL HRC", "HR=F", "HR=F", "HR=F", "HR=F", "HR=F"), _
        Array("S&P 500 E-MINI",  "ES=F",  "ES=F",  "ES=F",  "ES=F",  "ES=F"), _
        Array("NASDAQ-100 EMINI", "NQ=F", "NQ=F",  "NQ=F",  "NQ=F",  "NQ=F"), _
        Array("DOW E-MINI",      "YM=F",  "YM=F",  "YM=F",  "YM=F",  "YM=F"), _
        Array("RUSSELL 2000",    "RTY=F", "RTY=F", "RTY=F", "RTY=F", "RTY=F"), _
        Array("VIX",             "VX=F",  "VX=F",  "VX=F",  "VX=F",  "VX=F") _
    )
    
    Dim r As Integer, col As Integer
    For r = 0 To UBound(data)
        For col = 0 To 5
            wsTick.Cells(r + 2, col + 1).Value = data(r)(col)
        Next col
    Next r
    
    MsgBox "? TickerMap �w��s�I�@ " & UBound(data) + 1 & " �Ӱӫ~�C" & vbCr & _
           "�{�b���� RefreshDashboard", vbInformation
End Sub

' ================================================================
'  �ʺA�p��U��X���N�X + % Change �ץ�
'  ���N�쥻�� FixTickerMap �M RefreshDashboard ��������q��
' ================================================================

' �w�w ����N�X��Ӫ� �w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w
' Yahoo Finance ���f����GF=1�� G=2 H=3 J=4 K=5 M=6
'                         N=7�� Q=8 U=9 V=10 X=11 Z=12
Function MonthCode(m As Integer) As String
    Dim codes As String: codes = "FGHJKMNQUVXZ"
    MonthCode = Mid(codes, m, 1)
End Function

' �w�w �p��� n �Ӥ�᪺�X���N�X �w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w
Function FuturesTicker(baseSymbol As String, exchange As String, _
                       monthsAhead As Long) As String
    Dim targetDate As Date
    targetDate = DateAdd("m", monthsAhead, Date)
    
    Dim y As Integer: y = Year(targetDate)
    Dim m As Integer: m = Month(targetDate)
    
    ' �榡�GCLK26.NYM
    FuturesTicker = baseSymbol & MonthCode(m) & Right(CStr(y), 2) & exchange
End Function

' �w�w �U�ӫ~������ҫ�� �w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w
' .NYM = NYMEX  .CMX = COMEX  .CBT = CBOT
Function GetExchange(commodity As String) As String
    Select Case commodity
        Case "CL", "RB", "HO", "NG": GetExchange = ".NYM"
        Case "BZ": GetExchange = ".NYM"                        ' Brent �P NYMEX
        Case "GC", "SI", "HG", "ALI": GetExchange = ".CMX"
        Case "HR": GetExchange = ".CMX"
        Case "ZC", "ZS", "ZW": GetExchange = ".CBT"
        ' --- Equity index futures ---
        Case "ES", "NQ", "RTY": GetExchange = ".CME"           ' E-mini S&P / Nasdaq / Russell
        Case "YM": GetExchange = ".CBT"                        ' E-mini Dow on CBOT
        Case "VX": GetExchange = ".CFE"                        ' VIX on CFE
        Case Else: GetExchange = ".CMX"
    End Select
End Function

' Index futures only roll on quarterly cycle (H/M/U/Z = Mar/Jun/Sep/Dec)
Function QuarterlyMonthCode(m As Integer) As String
    Select Case m
        Case 1, 2, 3:    QuarterlyMonthCode = "H"   ' March
        Case 4, 5, 6:    QuarterlyMonthCode = "M"   ' June
        Case 7, 8, 9:    QuarterlyMonthCode = "U"   ' September
        Case 10, 11, 12: QuarterlyMonthCode = "Z"   ' December
    End Select
End Function

' Snap requested month-ahead forward to the NEXT quarterly contract.
Function FuturesTickerQuarterly(baseSymbol As String, exchange As String, _
                                 monthsAhead As Long) As String
    Dim targetDate As Date
    targetDate = DateAdd("m", monthsAhead, Date)
    Dim y As Integer: y = Year(targetDate)
    Dim m As Integer: m = Month(targetDate)
    FuturesTickerQuarterly = baseSymbol & QuarterlyMonthCode(m) & Right(CStr(y), 2) & exchange
End Function

' Returns True for symbols that only trade on the H/M/U/Z quarterly cycle.
Function IsQuarterlyOnly(baseSymbol As String) As Boolean
    Select Case baseSymbol
        Case "ES", "NQ", "YM", "RTY"
            IsQuarterlyOnly = True
        Case Else
            IsQuarterlyOnly = False
    End Select
End Function

' ================================================================
'  ���� TickerMap �X �ϥΰʺA����N�X
' ================================================================
Sub RebuildTickerMapDynamic()
    Dim wsTick As Worksheet
    On Error Resume Next
    Set wsTick = ThisWorkbook.Sheets("TickerMap")
    On Error GoTo 0
    If wsTick Is Nothing Then
        Set wsTick = ThisWorkbook.Sheets.Add
        wsTick.Name = "TickerMap"
        wsTick.Visible = xlSheetHidden
    End If
    wsTick.Cells.Clear
    
    ' ���D
    wsTick.Cells(1, 1) = "Instrument"
    wsTick.Cells(1, 2) = "Base"
    wsTick.Cells(1, 3) = "M1"
    wsTick.Cells(1, 4) = "M2"
    wsTick.Cells(1, 5) = "M3"
    wsTick.Cells(1, 6) = "M6"
    wsTick.Cells(1, 7) = "M12"
    
    ' [�ӫ~�W��, ��¦�N�X]
    Dim items As Variant
    items = Array( _
        Array("WTI CRUDE OIL", "CL"), _
        Array("BRENT CRUDE", "BZ"), _
        Array("RBOB GASOLINE", "RB"), _
        Array("ULSD DIESEL", "HO"), _
        Array("NATURAL GAS", "NG"), _
        Array("GOLD", "GC"), _
        Array("COPPER", "HG"), _
        Array("SILVER", "SI"), _
        Array("CORN", "ZC"), _
        Array("SOYBEANS", "ZS"), _
        Array("WHEAT", "ZW"), _
        Array("ALUMINUM", "ALI"), _
        Array("STEEL HRC", "HR") _
    )
    
    Dim offsets As Variant
    offsets = Array(1, 2, 3, 6, 12)  ' M1, M2, M3, M6, M12
    
    Dim i As Integer, j As Integer
    For i = 0 To UBound(items)
        Dim row As Integer: row = i + 2
        Dim base As String: base = items(i)(1)
        Dim exch As String: exch = GetExchange(base)
        
        wsTick.Cells(row, 1) = items(i)(0)
        wsTick.Cells(row, 2) = base
        
        For j = 0 To 4
            wsTick.Cells(row, j + 3) = FuturesTicker(base, exch, CInt(offsets(j)))
        Next j
    Next i
    
    ' ���G�w��
    Dim preview As String
    preview = "�ʺA�ͦ����X���N�X�]�H WTI ���ҡ^�G" & vbCr & vbCr
    For j = 0 To 4
        Dim mo As Variant: mo = Array("M1", "M2", "M3", "M6", "M12")
        preview = preview & mo(j) & " = " & wsTick.Cells(2, j + 3).Value & vbCr
    Next j
    
    MsgBox preview & vbCr & "�@ " & UBound(items) + 1 & " �Ӱӫ~�w��s�C", vbInformation
End Sub

' ================================================================
'  V3 TickerMap builder
'  - Uses Yahoo-style per-month tickers (<base><MonthCode><YY><.EXCH>)
'  - Honors quarterly-only contracts for equity index futures (ES/NQ/YM/RTY)
'  - Adds equity index futures (S&P 500, Nasdaq-100, Dow, Russell 2000, VIX)
'    forward curve alongside the commodity universe.
'  Layout (same as RebuildTickerMapDynamic):
'    A=Instrument | B=Base | C=M1 | D=M2 | E=M3 | F=M6 | G=M12
'  Use with RefreshDashboardV2.
' ================================================================
Sub RebuildTickerMapV3()
    Dim wsTick As Worksheet
    On Error Resume Next
    Set wsTick = ThisWorkbook.Sheets("TickerMap")
    On Error GoTo 0
    If wsTick Is Nothing Then
        Set wsTick = ThisWorkbook.Sheets.Add
        wsTick.Name = "TickerMap"
        wsTick.Visible = xlSheetHidden
    End If
    wsTick.Cells.Clear

    wsTick.Cells(1, 1) = "Instrument"
    wsTick.Cells(1, 2) = "Base"
    wsTick.Cells(1, 3) = "M1"
    wsTick.Cells(1, 4) = "M2"
    wsTick.Cells(1, 5) = "M3"
    wsTick.Cells(1, 6) = "M6"
    wsTick.Cells(1, 7) = "M12"

    ' [Instrument display name, Base symbol]
    Dim items As Variant
    items = Array( _
        Array("WTI CRUDE OIL",   "CL"), _
        Array("BRENT CRUDE",     "BZ"), _
        Array("RBOB GASOLINE",   "RB"), _
        Array("ULSD DIESEL",     "HO"), _
        Array("NATURAL GAS",     "NG"), _
        Array("GOLD",            "GC"), _
        Array("COPPER",          "HG"), _
        Array("SILVER",          "SI"), _
        Array("CORN",            "ZC"), _
        Array("SOYBEANS",        "ZS"), _
        Array("WHEAT",           "ZW"), _
        Array("ALUMINUM",        "ALI"), _
        Array("STEEL HRC",       "HR"), _
        Array("S&P 500 E-MINI",  "ES"), _
        Array("NASDAQ-100 EMINI", "NQ"), _
        Array("DOW E-MINI",      "YM"), _
        Array("RUSSELL 2000",    "RTY"), _
        Array("VIX",             "VX") _
    )

    Dim offsets As Variant
    offsets = Array(1, 2, 3, 6, 12)  ' M1, M2, M3, M6, M12

    Dim i As Integer, j As Integer
    For i = 0 To UBound(items)
        Dim r As Integer: r = i + 2
        Dim base As String: base = items(i)(1)
        Dim exch As String: exch = GetExchange(base)

        wsTick.Cells(r, 1) = items(i)(0)
        wsTick.Cells(r, 2) = base

        For j = 0 To 4
            If IsQuarterlyOnly(base) Then
                wsTick.Cells(r, j + 3) = FuturesTickerQuarterly(base, exch, CLng(offsets(j)))
            Else
                wsTick.Cells(r, j + 3) = FuturesTicker(base, exch, CLng(offsets(j)))
            End If
        Next j
    Next i

    MsgBox "TickerMap V3 rebuilt - " & UBound(items) + 1 & " instruments " & _
           "(commodities + equity index futures forward curve)." & vbCr & vbCr & _
           "WTI M1   = " & wsTick.Cells(2, 3).Value & vbCr & _
           "S&P M1   = " & wsTick.Cells(15, 3).Value & vbCr & _
           "S&P M6   = " & wsTick.Cells(15, 6).Value & vbCr & vbCr & _
           "Run SyncLayoutFromTickerMap then RefreshDashboard.", vbInformation
End Sub

' ================================================================
'  SyncLayoutFromTickerMap
'  - Reads every Instrument row from TickerMap and pushes it into
'    BBG_COMM Section 1 column B, then updates _Config!B2 (INST_COUNT).
'  - This is what makes newly added index futures actually appear in
'    the dashboard. Run once after RebuildTickerMapV3.
'  - If Section 1 is about to overflow into Section 2, shifts the
'    Section 2 anchor (_Config!B3) downward so they do not collide.
' ================================================================
Sub SyncLayoutFromTickerMap()
    Dim wsM    As Worksheet: Set wsM = ThisWorkbook.Sheets(SH_MAIN)
    Dim wsTick As Worksheet: Set wsTick = ThisWorkbook.Sheets(SH_TICK)
    Dim wsCfg  As Worksheet: Set wsCfg = ThisWorkbook.Sheets(SH_CFG)

    Dim sec1     As Long: sec1 = CLng(wsCfg.Range("B1").Value)
    Dim sec2Data As Long: sec2Data = CLng(wsCfg.Range("B3").Value)
    Dim oldCnt   As Long: oldCnt = CLng(wsCfg.Range("B2").Value)

    ' Count TickerMap data rows (skip header at row 1, stop on blank Instrument).
    Dim newCnt As Long: newCnt = 0
    Dim i As Long
    For i = 2 To 200
        If Trim(CStr(wsTick.Cells(i, 1).Value)) = "" Then Exit For
        newCnt = newCnt + 1
    Next i
    If newCnt = 0 Then
        MsgBox "TickerMap is empty - run RebuildTickerMapV3 first.", vbExclamation
        Exit Sub
    End If

    ' If the new instrument list runs past the existing Section 2 anchor,
    ' push Section 2 down (plus 1 blank gap row).
    If sec1 + newCnt > sec2Data Then
        Dim shift As Long: shift = (sec1 + newCnt + 1) - sec2Data
        sec2Data = sec2Data + shift
        wsCfg.Range("B3").Value = sec2Data
    End If

    ' Clear any rows beyond the new count (was 18 -> now 14 case).
    ' Unmerge each row first to survive previously merged title bands.
    If newCnt < oldCnt Then
        Dim wr As Long, cc As Integer
        For wr = sec1 + newCnt To sec1 + oldCnt - 1
            On Error Resume Next
            wsM.Range(wsM.Cells(wr, 1), wsM.Cells(wr, COL_LAST)).UnMerge
            On Error GoTo 0
            For cc = 1 To 13
                With wsM.Cells(wr, cc)
                    .ClearContents
                    .Interior.Color = C_BG
                    .Font.Color = C_BG
                    .Font.Name = THEME_FONT
                    .Font.Size = THEME_FONT_SIZE
                End With
            Next cc
        Next wr
    End If

    ' Write instrument names into BBG_COMM column B.
    For i = 1 To newCnt
        Dim r As Long: r = sec1 + i - 1
        Dim instName As String: instName = CStr(wsTick.Cells(i + 1, 1).Value)
        Dim baseSym  As String: baseSym = CStr(wsTick.Cells(i + 1, 2).Value)
        With wsM.Cells(r, 2)
            .Value = instName & "  [" & baseSym & "]"
            .Font.Name = THEME_FONT
            .Font.Size = THEME_FONT_SIZE
            .Font.Bold = True
            .Font.Color = RGB(220, 220, 220)
            .Interior.Color = C_BG
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
    Next i

    wsCfg.Range("B2").Value = newCnt

    MsgBox "Layout synced: " & newCnt & " instruments (was " & oldCnt & ")." & vbCr & _
           "Section 2 anchor: row " & sec2Data & vbCr & vbCr & _
           "Now run RefreshDashboard.", vbInformation
End Sub

' ================================================================
'  Cell-writing helpers
' ================================================================
Private Sub WritePrice(cell As Range, price As Double)
    cell.Interior.Color = C_BG
    cell.Font.Name = THEME_FONT
    cell.Font.Size = THEME_FONT_SIZE
    cell.Font.Bold = False
    cell.HorizontalAlignment = xlRight
    cell.VerticalAlignment = xlCenter
    If price > 0 Then
        cell.NumberFormat = "#,##0.00"
        cell.Value = price
        cell.Font.Color = RGB(220, 220, 220)
    Else
        ' No usable quote for this contract month - show a soft dash so
        ' the user can distinguish "not available" from a real zero price.
        cell.NumberFormat = "@"
        cell.Value = "-"
        cell.Font.Color = RGB(100, 100, 100)
        cell.HorizontalAlignment = xlCenter
    End If
End Sub

Private Sub WriteSpread(cell As Range, spd As Double)
    ' Show "0.00" for a genuine zero so e.g. equity-index rows (where M1..M12
    ' all equal the front-month spot) read as "calculated and equal to zero"
    ' rather than the dash format which implies missing data.
    cell.NumberFormat = "#,##0.00;(#,##0.00);0.00"
    cell.Value = spd
    cell.Font.Name = THEME_FONT
    cell.Font.Size = THEME_FONT_SIZE
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
    If spd > 0.005 Then
        cell.Font.Color = RGB(0, 255, 136)
        cell.Interior.Color = RGB(0, 40, 20)
    ElseIf spd < -0.005 Then
        cell.Font.Color = RGB(255, 51, 51)
        cell.Interior.Color = RGB(40, 0, 0)
    Else
        cell.Font.Color = RGB(255, 255, 0)
        cell.Interior.Color = RGB(40, 40, 0)
    End If
End Sub

Private Sub ApplyStructure(cell As Range, pM1 As Double, pM12 As Double)
    Dim spd As Double: spd = pM1 - pM12
    cell.Font.Name = THEME_FONT
    cell.Font.Size = THEME_FONT_SIZE
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
    If spd > 0.005 Then
        cell.Value = "BACKWARDATION"
        cell.Font.Color = RGB(0, 255, 136)
        cell.Interior.Color = RGB(0, 51, 25)
    ElseIf spd < -0.005 Then
        cell.Value = "CONTANGO"
        cell.Font.Color = RGB(255, 51, 51)
        cell.Interior.Color = RGB(51, 0, 0)
    Else
        cell.Value = "MIXED / FLAT"
        cell.Font.Color = RGB(255, 255, 0)
        cell.Interior.Color = RGB(40, 40, 0)
    End If
End Sub

' Only marks a row as ERR when *every* price column is empty/zero.
' If at least one month came back valid, the prior WritePrice calls already
' rendered the cells correctly and we should not overwrite them.
' Skips COL_UNUSED_GAP (col 12) which is reserved as a visual divider.
Private Sub MarkError(ws As Worksheet, r As Long)
    Dim c As Integer
    For c = 4 To COL_LAST
        If c = COL_UNUSED_GAP Then GoTo NextCol
        Dim v As Variant: v = ws.Cells(r, c).Value
        Dim hasData As Boolean: hasData = False
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then hasData = True
        End If
        If Not hasData Then
            With ws.Cells(r, c)
                .NumberFormat = "@"
                .Value = "ERR"
                .Font.Name = THEME_FONT
                .Font.Size = THEME_FONT_SIZE
                .Font.Bold = True
                .Font.Color = RGB(255, 102, 0)
                .Interior.Color = C_BG
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If
NextCol:
    Next c
    ' Make sure the gap column stays clean.
    With ws.Cells(r, COL_UNUSED_GAP)
        .ClearContents
        .Interior.Color = C_BG
    End With
End Sub

Private Function GetCellPrice(cell As Range) As Double
    If IsNumeric(cell.Value) Then GetCellPrice = CDbl(cell.Value) Else GetCellPrice = 0
End Function

' ================================================================
'  Utilities
' ================================================================
Private Sub PauseMS(ms As Long)
    Dim t As Single: t = Timer + ms / 1000
    Do While Timer < t: DoEvents: Loop
End Sub

Sub ExportSnapshot()
    Dim ws As Worksheet:  Set ws = ThisWorkbook.Sheets(SH_MAIN)
    Dim wsCfg As Worksheet: Set wsCfg = ThisWorkbook.Sheets(SH_CFG)

    ' Resolve a writable output directory.
    ' ThisWorkbook.Path is empty when the workbook has never been saved -
    ' that produced the Open-failure ("\TermStructure_*.csv" has no drive).
    Dim outDir As String
    If LenB(ThisWorkbook.Path) > 0 Then
        outDir = ThisWorkbook.Path
    Else
        outDir = Environ$("USERPROFILE") & "\Desktop"
        If Len(Dir(outDir, vbDirectory)) = 0 Then outDir = Environ$("TEMP")
    End If

    Dim outPath As String
    outPath = outDir & "\TermStructure_" & Format(Now(), "yyyymmdd_hhmmss") & ".csv"

    ' Derive the export row range from _Config instead of hard-coding 60.
    Dim sec1     As Long: sec1 = CLng(wsCfg.Range("B1").Value)
    Dim instCnt  As Long: instCnt = CLng(wsCfg.Range("B2").Value)
    Dim sec2Data As Long: sec2Data = CLng(wsCfg.Range("B3").Value)
    Dim lastRow  As Long: lastRow = sec2Data + 6   ' include timestamp row

    Dim fNum As Integer: fNum = FreeFile

    On Error GoTo BadOpen
    Open outPath For Output As #fNum
    On Error GoTo 0

    Dim r As Long, c As Integer, line As String
    For r = 1 To lastRow
        line = ""
        For c = 2 To COL_LAST
            Dim v As String: v = CStr(ws.Cells(r, c).Value)
            ' RFC 4180: wrap in quotes when the cell contains a comma,
            ' embedded quote, or newline; escape internal quotes by doubling.
            If InStr(v, ",") > 0 Or InStr(v, """") > 0 Or InStr(v, vbLf) > 0 Then
                v = """" & Replace(v, """", """""") & """"
            End If
            line = line & IIf(c > 2, ",", "") & v
        Next c
        Print #fNum, line
    Next r
    Close #fNum

    MsgBox "Exported:" & vbCr & outPath, vbInformation
    Exit Sub

BadOpen:
    Dim errDesc As String: errDesc = Err.Description
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    MsgBox "Cannot write CSV:" & vbCr & outPath & vbCr & vbCr & _
           "Error " & Err.Number & ": " & errDesc, vbCritical
End Sub
