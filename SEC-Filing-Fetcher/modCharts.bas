Attribute VB_Name = "modCharts"
Option Explicit

' Module-level declarations must all sit in this declarations section, before
' any Sub/Function/Property -- VBA's parser rejects a Const/Dim placed after a
' procedure has already been defined elsewhere in the file (a real compile
' error the first version of this line hit, sitting right before
' ExportTickerSnapshot further down: "只有註解可以放在 End Sub、End Function 或
' End Property 後面" is VBA's message for exactly that misplacement, not
' anything wrong with the Const statement itself).
'
' Hardcoded like build.ps1's own $savePath -- this repo's existing convention
' for this tool is a fixed Desktop-relative path, not something derived from
' ThisWorkbook.Path (which can come back as a https://d.docs.live.net/... URL
' instead of a filesystem path when OneDrive sync is involved, so deriving the
' export folder from it would be unreliable).
Private Const EXPORT_FOLDER As String = "C:\Users\ryan9\OneDrive\桌面\Claudecode\SEC-Filing-Fetcher\exports\"

' "No value" is represented as a zero-length string sentinel inside Variant
' arrays here. LookupConceptValue always returns its result typed as a String
' subtype Variant (even numeric results), so checking VarType against vbString
' can't distinguish "has a value" from "empty" -- go through CStr()/IsNumeric()
' instead, which also sidesteps the Type Mismatch a direct "v = ''" comparison
' would throw when v holds a Double.
Private Function HasVal(ByVal v As Variant) As Boolean
    HasVal = (Not IsNull(v)) And (CStr(v) <> "") And IsNumeric(v)
End Function

' Compound annual growth rate between two same-unit values n periods apart,
' e.g. (end/start)^(1/years) - 1. Blank unless both endpoints have a value AND
' are strictly positive -- a negative-to-positive (or negative-to-negative)
' span has no meaningful compound growth rate, and a fractional power of a
' negative base would just produce a nonsensical/complex result. Written as
' three separate early-exit checks rather than one "HasVal(x) And CDbl(x)>0"
' condition -- VBA's And does not short-circuit, so a combined condition would
' still evaluate CDbl on an empty-string sentinel and throw Type Mismatch
' (see the investedCapital comment further down this file for the same trap).
Private Function SafeCagr(ByVal startVal As Variant, ByVal endVal As Variant, ByVal years As Double) As Variant
    SafeCagr = ""
    If Not HasVal(startVal) Then Exit Function
    If Not HasVal(endVal) Then Exit Function
    If CDbl(startVal) <= 0 Then Exit Function
    If CDbl(endVal) <= 0 Then Exit Function
    SafeCagr = (CDbl(endVal) / CDbl(startVal)) ^ (1# / years) - 1
End Function

' Min/max across the numeric entries of arr(1..cnt), for fitting a chart's value
' axis to what it actually plots. Returns False when there's nothing numeric in
' there at all, in which case lo/hi are left alone.
Public Function ArrayMinMax(ByVal arr As Variant, ByVal cnt As Long, ByRef lo As Double, ByRef hi As Double) As Boolean
    Dim found As Boolean
    found = False

    Dim i As Long
    For i = 1 To cnt
        If HasVal(arr(i)) Then
            Dim d As Double
            d = CDbl(arr(i))
            If Not found Then
                lo = d
                hi = d
                found = True
            Else
                If d < lo Then lo = d
                If d > hi Then hi = d
            End If
        End If
    Next i

    ArrayMinMax = found
End Function

' Finds or creates the sheet but never clears it -- callers that need a clean
' slate must call ClearSheetForRebuild themselves, once, before writing. This
' split exists so multiple builders can safely share one sheet (Dashboard):
' each GetOrCreateSheet call used to wipe the whole sheet unconditionally,
' which would destroy an earlier builder's content on a shared sheet.
Public Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Sub ClearSheetForRebuild(ByVal ws As Worksheet)
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    ws.Cells.Clear
End Sub

' ---------- Per-ticker snapshot export ----------

' Called once a full fetch (US SEC or TW MOPS) finishes, so a single ticker's
' result can be handed off/archived without the whole multi-ticker working
' file. Saves a standalone, macro-free copy of just that ticker's
' Dashboard/RawData/QuarterlySnapshot (or the TW_-prefixed equivalents) into
' SEC-Filing-Fetcher/exports/, keeping only the newest export per ticker.
'
' wb.Sheets(Array(...)).Copy with NO destination argument is the deliberate
' choice over Workbooks.Add + per-sheet .Copy: Excel creates a brand-new
' workbook containing exactly those sheets, and because sheet NAMES are
' preserved 1:1 in the copy, every chart's SeriesCollection formula (which
' encodes a bare sheet name like "RawData!$A$3:$A$1252", not a workbook-
' qualified reference) keeps resolving correctly with no relinking step of our
' own. Copying sheets in one at a time risks a chart series left pointing back
' at the original, still-open workbook instead. Since the new workbook is
' built purely from worksheet copies, it never carries a VBProject at all --
' saving it as plain .xlsx (FileFormat 51, not the macro-enabled 52) needs no
' separate macro-stripping step, there's simply nothing to strip.
'
' On failure, shows its own MsgBox (the caller's existing ErrHandler only
' narrates errors into the status cell, which would make an export failure
' easy to miss) and then re-raises, so the failure aborts the whole fetch the
' same way any other error in FetchSECFilings/FetchMOPSFilings already does.
Public Sub ExportTickerSnapshot(ByVal wb As Workbook, ByVal ticker As String, _
    ByVal dashboardSheetName As String, ByVal rawDataSheetName As String, ByVal quarterlySheetName As String)

    Dim wasScreenUpdating As Boolean, wasDisplayAlerts As Boolean
    wasScreenUpdating = Application.ScreenUpdating
    wasDisplayAlerts = Application.DisplayAlerts

    On Error GoTo ExportErr

    If Dir(EXPORT_FOLDER, vbDirectory) = "" Then MkDir EXPORT_FOLDER

    Dim safeTicker As String
    safeTicker = SanitizeForFilename(ticker)

    ' Keep only the newest export per ticker -- delete any earlier
    ' <ticker>_*.xlsx before writing the new one rather than letting the
    ' folder accumulate one file per fetch forever.
    Dim oldFile As String
    oldFile = Dir(EXPORT_FOLDER & safeTicker & "_*.xlsx")
    Do While oldFile <> ""
        Kill EXPORT_FOLDER & oldFile
        oldFile = Dir()
    Loop

    Dim destPath As String
    destPath = EXPORT_FOLDER & safeTicker & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    wb.Sheets(Array(dashboardSheetName, rawDataSheetName, quarterlySheetName)).Copy
    Dim newWb As Workbook
    Set newWb = ActiveWorkbook
    newWb.SaveAs Filename:=destPath, FileFormat:=51   ' xlOpenXMLWorkbook (.xlsx, no macros)
    newWb.Close SaveChanges:=False

    Application.ScreenUpdating = wasScreenUpdating
    Application.DisplayAlerts = wasDisplayAlerts

    MsgBox "已匯出個股快照：" & destPath, vbInformation, "匯出完成"
    Exit Sub

ExportErr:
    Application.ScreenUpdating = wasScreenUpdating
    Application.DisplayAlerts = wasDisplayAlerts
    MsgBox "個股快照匯出失敗（" & ticker & "）：" & Err.Description, vbCritical, "匯出失敗"
    Err.Raise Err.Number, "ExportTickerSnapshot", Err.Description
End Sub

' Strips characters Windows filenames can't contain, so a ticker with an
' unexpected character (unlikely, but neither the US ticker box nor the TW
' CO_ID box validates input upstream) can't produce an invalid Dir/Kill/SaveAs
' path.
Private Function SanitizeForFilename(ByVal s As String) As String
    Dim bad As String
    bad = "\/:*?""<>|"
    Dim i As Long, result As String
    result = s
    For i = 1 To Len(bad)
        result = Replace(result, Mid$(bad, i, 1), "_")
    Next i
    SanitizeForFilename = result
End Function

' ---------- RawData sheet (every chart's source data, one series per column pair) ----------

' Dashboard holds only the snapshot table and the chart objects; the numbers
' behind those charts all live on the "RawData" sheet instead, laid out as
' adjacent 2-column blocks -- item 1 in columns A:B, item 2 in C:D, item 3 in
' E:F, and so on. Row 1 is the item name, row 2 the two column headers, rows 3+
' the data.
'
' Each pair carries its OWN key column (date / period label) rather than sharing
' one key column across the sheet: blocks have wildly different lengths (~1250
' daily price rows next to 7 annual values), and a self-contained rectangle per
' item is what lets each chart point at one contiguous range.
'
' col is advanced by 2 on success so callers can keep passing the same running
' variable; it is left alone when there's nothing to write, so a metric with no
' data doesn't leave an empty column pair mid-sheet. keyRange/valRange come back
' Nothing in that case -- callers must not chart from them.
Public Sub WriteRawBlock(ByVal wsRaw As Worksheet, ByRef col As Long, ByVal title As String, _
    ByVal keyHeader As String, ByVal valHeader As String, _
    ByVal keys As Variant, ByVal vals As Variant, ByVal cnt As Long, _
    ByVal keyFormat As String, ByVal valFormat As String, _
    ByRef keyRange As Range, ByRef valRange As Range)

    Set keyRange = Nothing
    Set valRange = Nothing
    If cnt <= 0 Then Exit Sub

    wsRaw.Cells(1, col).Value = title
    wsRaw.Cells(2, col).Value = keyHeader
    wsRaw.Cells(2, col + 1).Value = valHeader

    Set keyRange = wsRaw.Range(wsRaw.Cells(3, col), wsRaw.Cells(cnt + 2, col))
    Set valRange = wsRaw.Range(wsRaw.Cells(3, col + 1), wsRaw.Cells(cnt + 2, col + 1))

    ' Number formats are applied BEFORE the values land, not after: the period
    ' labels ("2025", "2026Q1") must be typed as text, and formatting a cell as
    ' text after a bare "2025" has already been written no longer un-numbers it.
    If keyFormat <> "" Then keyRange.NumberFormat = keyFormat
    If valFormat <> "" Then valRange.NumberFormat = valFormat

    Dim outArr() As Variant
    ReDim outArr(1 To cnt, 1 To 2)
    Dim i As Long
    For i = 1 To cnt
        outArr(i, 1) = keys(i)
        outArr(i, 2) = vals(i)
    Next i
    wsRaw.Range(wsRaw.Cells(3, col), wsRaw.Cells(cnt + 2, col + 1)).Value = outArr

    Dim blockRange As Range
    Set blockRange = wsRaw.Range(wsRaw.Cells(1, col), wsRaw.Cells(cnt + 2, col + 1))
    blockRange.Font.Color = RGB(255, 255, 255)
    blockRange.Font.Name = "Yu Gothic"
    Call ApplyGridBorder(blockRange)

    With wsRaw.Range(wsRaw.Cells(1, col), wsRaw.Cells(2, col + 1))
        .Font.Bold = True
        .Font.Color = RGB(255, 165, 0)
    End With

    ' Orange right edge marks where this item ends. The pairs sit flush against
    ' each other with no spacer column, so without a visible divider ~100
    ' columns of numbers read as one undifferentiated block.
    With blockRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(255, 165, 0)
    End With

    col = col + 2
End Sub

' Blanket black background well past the last used cell (both directions), so
' scrolling right across ~100 columns or down past the longest block doesn't hit
' a white cutoff -- same treatment every other sheet gets, just sized to a sheet
' that is far wider than it is deep.
Public Sub FinishRawDataSheet(ByVal wsRaw As Worksheet, ByVal nextFreeCol As Long)
    Dim usedRows As Long
    usedRows = wsRaw.UsedRange.Row + wsRaw.UsedRange.Rows.Count - 1

    Dim paintRows As Long, paintCols As Long
    paintRows = usedRows + 200
    If paintRows < 500 Then paintRows = 500
    paintCols = nextFreeCol + 10
    If paintCols < 60 Then paintCols = 60

    wsRaw.Range(wsRaw.Cells(1, 1), wsRaw.Cells(paintRows, paintCols)).Interior.Color = RGB(0, 0, 0)
    wsRaw.Columns.AutoFit
End Sub

' A freshly added ChartObject can arrive with Excel's own guess at a source
' range already plotted (it picks up whatever contiguous block the active cell
' happens to sit in). Every series below is added explicitly, so drop anything
' that came for free first -- otherwise a stray auto-series would ride along
' with its own color and legend entry.
Public Sub ClearAutoSeries(ByVal ch As Chart)
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
End Sub

' Writes a section-title row's text and records its row number into
' bannerRows; does NOT set colors here. Colors are applied in one final pass
' (see StyleBannerRows) after the table's blanket black-background paint runs
' -- applying them here would just get overwritten by that later paint, since
' this is called before the table's full extent (and hence its paint range)
' is known. Deliberately NOT paired with Freeze Panes (Dashboard is long, but
' the user wants visual landmarks to scroll/search for, not a locked header)
' -- don't add FreezePanes here.
Private Sub WriteSectionBanner(ByVal ws As Worksheet, ByRef curRow As Long, ByVal title As String, ByVal totalCols As Long, ByRef bannerRows As Collection)
    ws.Cells(curRow, 1).Value = title
    bannerRows.Add curRow
    curRow = curRow + 1
End Sub

' Applies the bold orange-on-dark-gray banner look to every row recorded by
' WriteSectionBanner, across columns A..totalCols. Called after the table's
' blanket black-background paint so the banner styling is what's left visible.
Private Sub StyleBannerRows(ByVal ws As Worksheet, ByVal bannerRows As Collection, ByVal totalCols As Long)
    Dim r As Variant
    For Each r In bannerRows
        With ws.Range(ws.Cells(CLng(r), 1), ws.Cells(CLng(r), totalCols))
            .Font.Bold = True
            .Font.Color = RGB(255, 165, 0)
            .Interior.Color = RGB(45, 45, 45)
        End With
    Next r
End Sub

' Writes one metric's label (col A) and its n per-fiscal-year values (cols
' B..n+1), formatted per fmt ("money" in millions / "price" 2dp / "ratio" 2dp /
' "pct" italic %), then advances curRow. Centralizes row bookkeeping so no
' caller has to hardcode row numbers in NumberFormat ranges.
Private Sub WriteMetricRow(ByVal ws As Worksheet, ByRef curRow As Long, ByVal label As String, ByVal dataArr As Variant, ByVal n As Long, ByVal fmt As String)
    ws.Cells(curRow, 1).Value = label
    Dim i As Long
    For i = 1 To n
        ws.Cells(curRow, i + 1).Value = dataArr(i)
    Next i

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, n + 1))
    Select Case fmt
        Case "money"
            rng.NumberFormat = "#,##0,,""M"""
        Case "price"
            rng.NumberFormat = "#,##0.00"
        Case "ratio"
            rng.NumberFormat = "0.00"
        Case "pct"
            rng.NumberFormat = "0.00%"
            rng.Font.Italic = True
    End Select

    curRow = curRow + 1
End Sub

' ---------- Dashboard orchestrator ----------

' Owns the single "Dashboard" sheet: clears it once, then stacks the snapshot
' table, the price/EPS chart, and the financial trend chart grid vertically
' (table width varies with how many fiscal years were fetched, so vertical
' stacking sidesteps needing a dynamic side-by-side width calculation). A
' banner row marks the start of each of the latter two blocks so a reader can
' visually locate them while scrolling a now much taller sheet.
'
' Dashboard deliberately carries NO chart source data: every series behind those
' charts is written to the separate "RawData" sheet (see WriteRawBlock) and the
' charts reference it from there. The ~1250-row daily price table used to sit
' between the snapshot table and the chart grid on this sheet, which pushed the
' trend charts thousands of pixels down and made the sheet unreadable by scroll.
' Both sheets are cleared here, once, before any builder writes to them.
Public Sub BuildDashboard(ByVal wb As Workbook, ByVal ticker As String, ByVal entityName As String, _
    ByVal filings10K As Collection, ByVal filings10Q As Collection, ByVal mapEps As Object, ByVal prices As Object, _
    ByVal mapRevenue As Object, ByVal mapShares As Object, ByVal allMaps As Variant, _
    ByVal mapInventory As Object, ByVal mapAR As Object, ByVal mapCurrentAssets As Object, ByVal mapCurrentLiabilities As Object, _
    ByVal mapLongTermDebt As Object, ByVal mapStockholdersEquity As Object, ByVal mapEffectiveTaxRate As Object, ByVal mapCapEx As Object, ByVal mapCFO As Object, _
    ByVal mapCash As Object, ByVal mapDA As Object, ByVal mapOperatingIncome As Object, ByVal mapDividends As Object, ByVal mapNetIncome As Object, _
    ByVal mapShortTermDebt As Object, ByVal mapCOGS As Object, ByVal mapAccountsPayable As Object, ByVal mapAssets As Object, ByVal filingsLastRow As Long, _
    ByVal wsFilings As Worksheet, _
    Optional ByVal dashboardSheetName As String = "Dashboard", Optional ByVal rawDataSheetName As String = "RawData", _
    Optional ByVal annualFormMarker As String = "10-K")

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, dashboardSheetName)
    Call ClearSheetForRebuild(ws)

    Dim wsRaw As Worksheet
    Set wsRaw = GetOrCreateSheet(wb, rawDataSheetName)
    Call ClearSheetForRebuild(wsRaw)

    ' Running left edge of the next free column pair on RawData. Order is fixed:
    ' price, then EPS, then each Filings metric with its annual block
    ' immediately followed by its quarterly one.
    Dim rawCol As Long
    rawCol = 1

    Dim tableLastRow As Long, tableLastCol As Long
    Call BuildSnapshotTableInto(ws, entityName, ticker, filings10K, mapRevenue, mapEps, mapShares, allMaps, prices, _
        mapInventory, mapAR, mapCurrentAssets, mapCurrentLiabilities, mapLongTermDebt, mapStockholdersEquity, mapEffectiveTaxRate, mapCapEx, mapCFO, _
        mapCash, mapDA, mapOperatingIncome, mapDividends, mapNetIncome, mapShortTermDebt, mapCOGS, mapAccountsPayable, mapAssets, False, tableLastRow, tableLastCol)

    ' Both banners are styled at the very end, after every chart builder has
    ' returned -- BuildFinancialChartsInto does a blanket black-background
    ' repaint starting from row 1, which silently wipes out gray banner styling
    ' applied any earlier than that (it used to eat the price banner's, which is
    ' why that one rendered as bare orange-on-black rather than on gray).
    Dim priceBannerRow As Long
    priceBannerRow = tableLastRow + 2
    ws.Cells(priceBannerRow, 1).Value = "═══ 股價走勢 Price History（原始資料：RawData 工作表）═══"

    Dim priceChartTop As Double, priceChartBottom As Double
    priceChartTop = ws.Cells(priceBannerRow + 1, 1).Top
    Dim priceChartBottomOut As Double
    Call BuildPriceChartInto(ws, wsRaw, rawCol, ticker, filings10K, filings10Q, mapEps, prices, priceBannerRow + 1, 300, priceChartTop, priceChartBottomOut)
    priceChartBottom = priceChartBottomOut

    ' The financial-charts grid is placed by pixel Top/Left, not by row, so its
    ' banner's row is only an estimate (default ~15pt row height) -- close
    ' enough to sit visually just above the grid, not pixel-exact.
    Dim chartsBannerRow As Long
    chartsBannerRow = CLng((priceChartBottom + 10) / 15)
    ws.Cells(chartsBannerRow, 1).Value = "═══ 財務趨勢圖 Financial Trend Charts（原始資料：RawData 工作表）═══"

    Call BuildFinancialChartsInto(ws, wsFilings, filingsLastRow, wsRaw, rawCol, 10, priceChartBottom + 30, annualFormMarker)

    Call StyleOneBannerRow(ws, priceBannerRow, tableLastCol)
    Call StyleOneBannerRow(ws, chartsBannerRow, tableLastCol)

    ws.Columns.AutoFit
    Call FinishRawDataSheet(wsRaw, rawCol)
End Sub

' Owns the single "QuarterlySnapshot" sheet: the same snapshot-table layout as
' Dashboard's annual table (same sections, same metrics, same formulas -- see
' BuildSnapshotTableInto), just fed filings10Q instead of filings10K so every
' column is a fiscal quarter instead of a fiscal year. No charts here and no
' RawData writes -- the quarterly trend charts already live on Dashboard
' (BuildFinancialChartsInto's "季度 Quarterly" series), this sheet is purely
' the tabular quarter-by-quarter read.
Public Sub BuildQuarterlyDashboard(ByVal wb As Workbook, ByVal ticker As String, ByVal entityName As String, _
    ByVal filings10Q As Collection, ByVal mapRevenue As Object, ByVal mapEps As Object, ByVal mapShares As Object, ByVal allMaps As Variant, ByVal prices As Object, _
    ByVal mapInventory As Object, ByVal mapAR As Object, ByVal mapCurrentAssets As Object, ByVal mapCurrentLiabilities As Object, _
    ByVal mapLongTermDebt As Object, ByVal mapStockholdersEquity As Object, ByVal mapEffectiveTaxRate As Object, ByVal mapCapEx As Object, ByVal mapCFO As Object, _
    ByVal mapCash As Object, ByVal mapDA As Object, ByVal mapOperatingIncome As Object, ByVal mapDividends As Object, ByVal mapNetIncome As Object, _
    ByVal mapShortTermDebt As Object, ByVal mapCOGS As Object, ByVal mapAccountsPayable As Object, ByVal mapAssets As Object, Optional ByVal sheetName As String = "QuarterlySnapshot")

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, sheetName)
    Call ClearSheetForRebuild(ws)

    Dim tableLastRow As Long, tableLastCol As Long
    Call BuildSnapshotTableInto(ws, entityName, ticker, filings10Q, mapRevenue, mapEps, mapShares, allMaps, prices, _
        mapInventory, mapAR, mapCurrentAssets, mapCurrentLiabilities, mapLongTermDebt, mapStockholdersEquity, mapEffectiveTaxRate, mapCapEx, mapCFO, _
        mapCash, mapDA, mapOperatingIncome, mapDividends, mapNetIncome, mapShortTermDebt, mapCOGS, mapAccountsPayable, mapAssets, True, tableLastRow, tableLastCol)

    ws.Columns.AutoFit
End Sub

' "Nothing to chart here" message, in place of the chart that would have been.
' Explicitly white-on-black: the surrounding sheet is painted black, so a cell
' written afterwards with default font color is black text on black.
Private Sub WriteDashboardNote(ByVal ws As Worksheet, ByVal r As Long, ByVal msg As String)
    With ws.Cells(r, 1)
        .Value = msg
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Yu Gothic"
    End With
End Sub

Private Sub StyleOneBannerRow(ByVal ws As Worksheet, ByVal r As Long, ByVal totalCols As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, totalCols))
        .Font.Bold = True
        .Font.Color = RGB(255, 165, 0)
        .Interior.Color = RGB(45, 45, 45)
    End With
End Sub

' Chart 1: last ~5 years of daily closing price plotted against each filing's
' reported diluted EPS, held flat until the next filing supersedes it (a step
' function), so the reader can see price movement against what was actually
' known/reported at each point in time. No Forward P/E series.
' The two series are written to RawData as two column pairs (date+price, then
' date+EPS) starting at rawCol, which is advanced past them. noteRow is only
' used to place a "no data" message on the Dashboard when there's nothing to
' chart. chartLeft/chartTop: pixel position of the floating chart object on the
' Dashboard. chartBottomOut reports the chart's pixel bottom edge so the caller
' can stack the next block below it.
Private Sub BuildPriceChartInto(ByVal ws As Worksheet, ByVal wsRaw As Worksheet, ByRef rawCol As Long, ByVal ticker As String, ByVal filings10K As Collection, ByVal filings10Q As Collection, _
    ByVal mapEps As Object, ByVal prices As Object, ByVal noteRow As Long, ByVal chartLeft As Double, ByVal chartTop As Double, ByRef chartBottomOut As Double)

    chartBottomOut = chartTop

    Dim n As Long
    n = filings10K.Count + filings10Q.Count
    If n = 0 Then Exit Sub

    Dim fDate() As String, fEps() As Double, fHasEps() As Boolean
    ReDim fDate(1 To n)
    ReDim fEps(1 To n)
    ReDim fHasEps(1 To n)

    Dim idx As Long
    idx = 0
    Dim f As Variant
    Dim v As Variant
    For Each f In filings10K
        idx = idx + 1
        fDate(idx) = f("filingDate")
        v = LookupConceptValue(mapEps, f("accn"), f("reportDate"), f("form"))
        If HasVal(v) Then
            fEps(idx) = CDbl(v)
            fHasEps(idx) = True
        End If
    Next f
    For Each f In filings10Q
        idx = idx + 1
        fDate(idx) = f("filingDate")
        v = LookupConceptValue(mapEps, f("accn"), f("reportDate"), f("form"))
        If HasVal(v) Then
            fEps(idx) = CDbl(v)
            fHasEps(idx) = True
        End If
    Next f

    Dim i As Long, j As Long
    For i = 1 To n - 1
        For j = 1 To n - i
            If fDate(j) > fDate(j + 1) Then
                Dim tS As String, tD As Double, tB As Boolean
                tS = fDate(j): fDate(j) = fDate(j + 1): fDate(j + 1) = tS
                tD = fEps(j): fEps(j) = fEps(j + 1): fEps(j + 1) = tD
                tB = fHasEps(j): fHasEps(j) = fHasEps(j + 1): fHasEps(j + 1) = tB
            End If
        Next j
    Next i

    Dim todayD As Date
    todayD = Date

    If prices.Count = 0 Then
        Call WriteDashboardNote(ws, noteRow, "找不到 " & ticker & " 的股價資料（Yahoo Finance 無資料，可能不是美股代號）")
        Exit Sub
    End If

    Dim dateKeys() As String
    ReDim dateKeys(1 To prices.Count)
    Dim kk As Long
    kk = 0
    Dim k As Variant
    For Each k In prices.Keys
        kk = kk + 1
        dateKeys(kk) = CStr(k)
    Next k

    Dim a As Long, b As Long, keyVal As String
    For a = 2 To UBound(dateKeys)
        keyVal = dateKeys(a)
        b = a - 1
        Do While b >= 1 And dateKeys(b) > keyVal
            dateKeys(b + 1) = dateKeys(b)
            b = b - 1
        Loop
        dateKeys(b + 1) = keyVal
    Next a

    Dim windowStartStr As String
    windowStartStr = Format$(DateAdd("yyyy", -5, todayD), "yyyy-mm-dd")

    Dim dateArr() As Variant, priceArr() As Variant, epsArr() As Variant
    ReDim dateArr(1 To prices.Count)
    ReDim priceArr(1 To prices.Count)
    ReDim epsArr(1 To prices.Count)

    Dim cnt As Long
    cnt = 0
    For kk = 1 To UBound(dateKeys)
        If dateKeys(kk) >= windowStartStr Then
            Dim curDateStr As String
            curDateStr = dateKeys(kk)

            Dim stepEps As Variant
            stepEps = ""
            For i = 1 To n
                If fDate(i) <= curDateStr And fHasEps(i) Then
                    stepEps = fEps(i)
                End If
            Next i

            cnt = cnt + 1
            dateArr(cnt) = CDate(curDateStr)
            priceArr(cnt) = prices(curDateStr)
            ' Empty, not 0 or "": dates before the first filing with an EPS must
            ' land in genuinely blank cells so the area series starts where the
            ' data does instead of hugging zero for the first stretch.
            If HasVal(stepEps) Then epsArr(cnt) = CDbl(stepEps) Else epsArr(cnt) = Empty
        End If
    Next kk

    If cnt = 0 Then
        Call WriteDashboardNote(ws, noteRow, "近5年無股價資料")
        Exit Sub
    End If

    Dim priceKeyRng As Range, priceValRng As Range
    Dim epsKeyRng As Range, epsValRng As Range
    Call WriteRawBlock(wsRaw, rawCol, ticker & " 股價（近5年日收盤）", "日期", ticker & " 股價", _
        dateArr, priceArr, cnt, "yyyy-mm-dd", "#,##0.00", priceKeyRng, priceValRng)
    Call WriteRawBlock(wsRaw, rawCol, "Diluted EPS (as reported)（申報後步進值）", "日期", "Diluted EPS", _
        dateArr, epsArr, cnt, "yyyy-mm-dd", "0.00", epsKeyRng, epsValRng)

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=750, Height:=400)
    Dim ch As Chart
    Set ch = co.Chart
    Call ClearAutoSeries(ch)
    ch.ChartType = xlLine

    ' Added EPS-first so it stays SeriesCollection(1) -- the area series has to
    ' be drawn behind the price line, and ApplyDarkThemeToChart colors series by
    ' index (dark orange fill, bright orange line).
    Dim serEps As Series, serPrice As Series
    Set serEps = ch.SeriesCollection.NewSeries
    serEps.Values = epsValRng
    serEps.XValues = epsKeyRng
    serEps.Name = "Diluted EPS (as reported)"
    serEps.ChartType = xlArea

    Set serPrice = ch.SeriesCollection.NewSeries
    serPrice.Values = priceValRng
    serPrice.XValues = priceKeyRng
    serPrice.Name = ticker & " Price"
    serPrice.ChartType = xlLine
    serPrice.AxisGroup = xlSecondary
    serPrice.Format.Line.Weight = 0.5

    ch.HasTitle = True
    ch.ChartTitle.Text = ticker & " — 近5年股價 vs. 已申報每股盈餘"

    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionTop

    ' Category axis holds real date values (not text), so switch it to a true
    ' time-scale axis and force one tick per calendar year -- otherwise Excel
    ' auto-picks a label every ~2 months across 5 years of daily data.
    Dim axDate As Axis
    Set axDate = ch.Axes(xlCategory)
    axDate.CategoryType = xlTimeScale
    axDate.BaseUnit = xlDays
    axDate.MajorUnit = 1
    axDate.MajorUnitScale = xlYears

    Call ApplyDarkThemeToChart(ch, Array(RGB(180, 90, 0), RGB(255, 180, 60)))

    ' Fit both value axes to their own series, after the theme pass (which reads
    ' back whatever Excel auto-picked), then re-fit whichever ended up with
    ' fewer major units so the two axes are divided the same number of times and
    ' their gridlines land on the same horizontal lines.
    Dim epsLo As Double, epsHi As Double, priceLo As Double, priceHi As Double
    Dim epsUnits As Long, priceUnits As Long
    epsUnits = 0
    priceUnits = 0
    If ArrayMinMax(epsArr, cnt, epsLo, epsHi) Then epsUnits = FitValueAxis(ch.Axes(xlValue, xlPrimary), epsLo, epsHi)
    If ArrayMinMax(priceArr, cnt, priceLo, priceHi) Then priceUnits = FitValueAxis(ch.Axes(xlValue, xlSecondary), priceLo, priceHi)
    If epsUnits > 0 And priceUnits > 0 And epsUnits <> priceUnits Then
        If epsUnits < priceUnits Then
            Call FitValueAxis(ch.Axes(xlValue, xlPrimary), epsLo, epsHi, priceUnits)
        Else
            Call FitValueAxis(ch.Axes(xlValue, xlSecondary), priceLo, priceHi, epsUnits)
        End If
    End If

    chartBottomOut = co.Top + co.Height
End Sub

' Table: actual historical fiscal-year snapshot (Stock Price / Market Cap /
' Revenue / EPS / P-E / P-S / growth rates / valuation & return ratios) across
' the 10-Ks already fetched, grouped into labeled sections (market data / income
' / valuation multiples / profitability & returns / working capital / capital
' structure / cash flow / dividends), plus a trailing P/E and P/S percentile
' band. Forecast years are intentionally excluded -- SEC filings only report
' what has actually happened. tableLastRow/tableLastCol report the bottom-right
' cell actually used so the caller can stack the next block below it.
' isQuarterly only changes cosmetics (column-header format, the percentile
' band's "近N年/季" label) -- the metric computation itself is period-agnostic,
' it just uses whatever accn/reportDate/form each passed-in filing carries, so
' the same code serves both the annual (10-K) and quarterly (10-Q) callers.
Private Sub BuildSnapshotTableInto(ByVal ws As Worksheet, ByVal entityName As String, ByVal ticker As String, ByVal filings10K As Collection, _
    ByVal mapRevenue As Object, ByVal mapEps As Object, ByVal mapShares As Object, ByVal allMaps As Variant, ByVal prices As Object, _
    ByVal mapInventory As Object, ByVal mapAR As Object, ByVal mapCurrentAssets As Object, ByVal mapCurrentLiabilities As Object, _
    ByVal mapLongTermDebt As Object, ByVal mapStockholdersEquity As Object, ByVal mapEffectiveTaxRate As Object, ByVal mapCapEx As Object, ByVal mapCFO As Object, _
    ByVal mapCash As Object, ByVal mapDA As Object, ByVal mapOperatingIncome As Object, ByVal mapDividends As Object, ByVal mapNetIncome As Object, _
    ByVal mapShortTermDebt As Object, ByVal mapCOGS As Object, ByVal mapAccountsPayable As Object, ByVal mapAssets As Object, ByVal isQuarterly As Boolean, _
    ByRef tableLastRow As Long, ByRef tableLastCol As Long)

    On Error GoTo BSTIErr

    Dim accn As String, reportDate As String, form As String
    Dim i As Long

    tableLastRow = 1
    tableLastCol = 1

    If filings10K.Count = 0 Then
        ws.Range("A1").Value = "沒有可用的 10-K 資料"
        tableLastRow = 1
        Exit Sub
    End If

    Dim n As Long
    n = filings10K.Count
    Dim ordered() As Object
    ReDim ordered(1 To n)
    i = 0
    Dim f As Variant
    For Each f In filings10K
        i = i + 1
        Set ordered(n - i + 1) = f
    Next f

    ws.Range("A1").Value = entityName & " (" & ticker & ")"
    ws.Range("A1").Font.Bold = True

    Dim priceRow() As Variant, mktCapRow() As Variant, sharesRow() As Variant, revRow() As Variant
    Dim epsRow() As Variant, salesGrowthRow() As Variant, epsGrowthRow() As Variant
    Dim peRow() As Variant, psRow() As Variant, spsRow() As Variant
    Dim invRow() As Variant, arRow() As Variant, curAssetsRow() As Variant, curLiabRow() As Variant, curRatioRow() As Variant
    Dim ltDebtRow() As Variant, stDebtRow() As Variant, equityRow() As Variant, bookValRow() As Variant, taxRateRow() As Variant, capexRow() As Variant, fcfRow() As Variant
    Dim cashRow() As Variant, opIncRow() As Variant, daRow() As Variant, evRow() As Variant, ebitdaRow() As Variant
    Dim evEbitdaRow() As Variant, evSalesRow() As Variant, roeRow() As Variant, roicRow() As Variant
    Dim dpsRow() As Variant, netIncRow() As Variant, divYieldRow() As Variant, payoutRow() As Variant, pegRow() As Variant
    ' ---- DuPont / CCC / FCF ratios / CAGR (added for the fundamental-analysis
    ' expansion) -- see the computation block after the per-period loop below.
    Dim cogsRow() As Variant, apRow() As Variant, assetsRow() As Variant
    Dim netMarginRow() As Variant, assetTurnoverRow() As Variant, equityMultiplierRow() As Variant
    Dim dioRow() As Variant, dsoRow() As Variant, dpoRow() As Variant, cccRow() As Variant
    Dim fcfConversionRow() As Variant, fcfYieldRow() As Variant
    Dim revCagr3Row() As Variant, epsCagr3Row() As Variant, fcfCagr3Row() As Variant
    Dim revCagr5Row() As Variant, epsCagr5Row() As Variant, fcfCagr5Row() As Variant
    ReDim priceRow(1 To n): ReDim mktCapRow(1 To n): ReDim sharesRow(1 To n): ReDim revRow(1 To n)
    ReDim epsRow(1 To n): ReDim salesGrowthRow(1 To n): ReDim epsGrowthRow(1 To n)
    ReDim peRow(1 To n): ReDim psRow(1 To n): ReDim spsRow(1 To n)
    ReDim invRow(1 To n): ReDim arRow(1 To n): ReDim curAssetsRow(1 To n): ReDim curLiabRow(1 To n): ReDim curRatioRow(1 To n)
    ReDim ltDebtRow(1 To n): ReDim stDebtRow(1 To n): ReDim equityRow(1 To n): ReDim bookValRow(1 To n): ReDim taxRateRow(1 To n): ReDim capexRow(1 To n): ReDim fcfRow(1 To n)
    ReDim cashRow(1 To n): ReDim opIncRow(1 To n): ReDim daRow(1 To n): ReDim evRow(1 To n): ReDim ebitdaRow(1 To n)
    ReDim evEbitdaRow(1 To n): ReDim evSalesRow(1 To n): ReDim roeRow(1 To n): ReDim roicRow(1 To n)
    ReDim dpsRow(1 To n): ReDim netIncRow(1 To n): ReDim divYieldRow(1 To n): ReDim payoutRow(1 To n): ReDim pegRow(1 To n)
    ReDim cogsRow(1 To n): ReDim apRow(1 To n): ReDim assetsRow(1 To n)
    ReDim netMarginRow(1 To n): ReDim assetTurnoverRow(1 To n): ReDim equityMultiplierRow(1 To n)
    ReDim dioRow(1 To n): ReDim dsoRow(1 To n): ReDim dpoRow(1 To n): ReDim cccRow(1 To n)
    ReDim fcfConversionRow(1 To n): ReDim fcfYieldRow(1 To n)
    ReDim revCagr3Row(1 To n): ReDim epsCagr3Row(1 To n): ReDim fcfCagr3Row(1 To n)
    ReDim revCagr5Row(1 To n): ReDim epsCagr5Row(1 To n): ReDim fcfCagr5Row(1 To n)

    For i = 1 To n
        accn = ordered(i)("accn")
        reportDate = ordered(i)("reportDate")
        form = ordered(i)("form")

        Dim fyfp As String, parts As Variant, colHeader As String
        fyfp = LookupFyFp(allMaps, accn, reportDate, form)
        parts = Split(fyfp, "|")
        If parts(0) <> "" Then
            If isQuarterly And parts(1) <> "" Then
                colHeader = "FY" & parts(0) & parts(1)
            Else
                colHeader = "FY" & parts(0)
            End If
        ElseIf isQuarterly Then
            colHeader = Format$(CDate(reportDate), "yyyy-mm")
        Else
            colHeader = Format$(CDate(reportDate), "yyyy")
        End If
        ws.Cells(1, i + 1).Value = colHeader
        ws.Cells(1, i + 1).Font.Bold = True

        Dim priceV As Variant, sharesV As Variant, revV As Variant, epsV As Variant
        priceV = NearestPriceOnOrBefore(prices, reportDate)
        sharesV = LookupConceptValue(mapShares, accn, reportDate, form)
        revV = LookupConceptValue(mapRevenue, accn, reportDate, form)
        epsV = LookupConceptValue(mapEps, accn, reportDate, form)

        If HasVal(priceV) Then priceRow(i) = CDbl(priceV) Else priceRow(i) = ""
        If HasVal(sharesV) Then sharesRow(i) = CDbl(sharesV) Else sharesRow(i) = ""
        If HasVal(revV) Then revRow(i) = CDbl(revV) Else revRow(i) = ""
        If HasVal(epsV) Then epsRow(i) = CDbl(epsV) Else epsRow(i) = ""

        If HasVal(priceRow(i)) And HasVal(sharesRow(i)) Then
            mktCapRow(i) = priceRow(i) * sharesRow(i)
        Else
            mktCapRow(i) = ""
        End If

        If HasVal(revRow(i)) And HasVal(sharesRow(i)) And sharesRow(i) <> 0 Then
            spsRow(i) = revRow(i) / sharesRow(i)
        Else
            spsRow(i) = ""
        End If

        If HasVal(priceRow(i)) And HasVal(epsRow(i)) And epsRow(i) <> 0 Then
            peRow(i) = priceRow(i) / epsRow(i)
        Else
            peRow(i) = ""
        End If

        If HasVal(mktCapRow(i)) And HasVal(revRow(i)) And revRow(i) <> 0 Then
            psRow(i) = mktCapRow(i) / revRow(i)
        Else
            psRow(i) = ""
        End If

        If i > 1 Then
            If HasVal(revRow(i)) And HasVal(revRow(i - 1)) And revRow(i - 1) <> 0 Then
                salesGrowthRow(i) = revRow(i) / revRow(i - 1) - 1
            Else
                salesGrowthRow(i) = ""
            End If
            If HasVal(epsRow(i)) And HasVal(epsRow(i - 1)) And epsRow(i - 1) <> 0 Then
                epsGrowthRow(i) = epsRow(i) / epsRow(i - 1) - 1
            Else
                epsGrowthRow(i) = ""
            End If
        Else
            salesGrowthRow(i) = ""
            epsGrowthRow(i) = ""
        End If

        Dim invV As Variant, arV As Variant, curAssetsV As Variant, curLiabV As Variant
        Dim ltDebtV As Variant, stDebtV As Variant, equityV As Variant, taxRateV As Variant, capexV As Variant, cfoV As Variant
        invV = LookupConceptValue(mapInventory, accn, reportDate, form)
        arV = LookupConceptValue(mapAR, accn, reportDate, form)
        curAssetsV = LookupConceptValue(mapCurrentAssets, accn, reportDate, form)
        curLiabV = LookupConceptValue(mapCurrentLiabilities, accn, reportDate, form)
        ltDebtV = LookupConceptValue(mapLongTermDebt, accn, reportDate, form)
        stDebtV = LookupConceptValue(mapShortTermDebt, accn, reportDate, form)
        equityV = LookupConceptValue(mapStockholdersEquity, accn, reportDate, form)
        taxRateV = LookupConceptValue(mapEffectiveTaxRate, accn, reportDate, form)
        capexV = LookupConceptValue(mapCapEx, accn, reportDate, form)
        cfoV = LookupConceptValue(mapCFO, accn, reportDate, form)

        Dim cogsV2 As Variant, apV As Variant, assetsV2 As Variant
        cogsV2 = LookupConceptValue(mapCOGS, accn, reportDate, form)
        apV = LookupConceptValue(mapAccountsPayable, accn, reportDate, form)
        assetsV2 = LookupConceptValue(mapAssets, accn, reportDate, form)
        If HasVal(cogsV2) Then cogsRow(i) = CDbl(cogsV2) Else cogsRow(i) = ""
        If HasVal(apV) Then apRow(i) = CDbl(apV) Else apRow(i) = ""
        If HasVal(assetsV2) Then assetsRow(i) = CDbl(assetsV2) Else assetsRow(i) = ""

        If HasVal(invV) Then invRow(i) = CDbl(invV) Else invRow(i) = ""
        If HasVal(arV) Then arRow(i) = CDbl(arV) Else arRow(i) = ""
        If HasVal(curAssetsV) Then curAssetsRow(i) = CDbl(curAssetsV) Else curAssetsRow(i) = ""
        If HasVal(curLiabV) Then curLiabRow(i) = CDbl(curLiabV) Else curLiabRow(i) = ""
        If HasVal(ltDebtV) Then ltDebtRow(i) = CDbl(ltDebtV) Else ltDebtRow(i) = ""
        If HasVal(stDebtV) Then stDebtRow(i) = CDbl(stDebtV) Else stDebtRow(i) = ""
        If HasVal(equityV) Then equityRow(i) = CDbl(equityV) Else equityRow(i) = ""
        If HasVal(taxRateV) Then taxRateRow(i) = CDbl(taxRateV) Else taxRateRow(i) = ""
        If HasVal(capexV) Then capexRow(i) = CDbl(capexV) Else capexRow(i) = ""

        If HasVal(curAssetsRow(i)) And HasVal(curLiabRow(i)) And curLiabRow(i) <> 0 Then
            curRatioRow(i) = curAssetsRow(i) / curLiabRow(i)
        Else
            curRatioRow(i) = ""
        End If

        If HasVal(equityRow(i)) And HasVal(sharesRow(i)) And sharesRow(i) <> 0 Then
            bookValRow(i) = equityRow(i) / sharesRow(i)
        Else
            bookValRow(i) = ""
        End If

        If HasVal(cfoV) And HasVal(capexRow(i)) Then
            fcfRow(i) = CDbl(cfoV) - capexRow(i)
        Else
            fcfRow(i) = ""
        End If

        ' ---- Valuation & return ratios ----
        Dim cashV As Variant, opIncV As Variant, daV As Variant, dpsV As Variant, netIncV As Variant
        cashV = LookupConceptValue(mapCash, accn, reportDate, form)
        opIncV = LookupConceptValue(mapOperatingIncome, accn, reportDate, form)
        daV = LookupConceptValue(mapDA, accn, reportDate, form)
        dpsV = LookupConceptValue(mapDividends, accn, reportDate, form)
        netIncV = LookupConceptValue(mapNetIncome, accn, reportDate, form)

        If HasVal(cashV) Then cashRow(i) = CDbl(cashV) Else cashRow(i) = ""
        If HasVal(opIncV) Then opIncRow(i) = CDbl(opIncV) Else opIncRow(i) = ""
        If HasVal(daV) Then daRow(i) = CDbl(daV) Else daRow(i) = ""
        If HasVal(dpsV) Then dpsRow(i) = CDbl(dpsV) Else dpsRow(i) = ""
        If HasVal(netIncV) Then netIncRow(i) = CDbl(netIncV) Else netIncRow(i) = ""

        ' Total debt = LT debt + short-term/current debt. A missing short-term
        ' figure is treated as 0 (not "unknown") -- most filers without the tag
        ' genuinely carry no short-term borrowings; LT debt missing still blanks
        ' the whole EV, since that's the dominant term for most companies.
        Dim totalDebtForEV As Variant
        If HasVal(ltDebtRow(i)) Then
            totalDebtForEV = ltDebtRow(i)
            If HasVal(stDebtRow(i)) Then totalDebtForEV = totalDebtForEV + stDebtRow(i)
        Else
            totalDebtForEV = ""
        End If

        If HasVal(mktCapRow(i)) And HasVal(totalDebtForEV) And HasVal(cashRow(i)) Then
            evRow(i) = mktCapRow(i) + totalDebtForEV - cashRow(i)
        Else
            evRow(i) = ""
        End If

        If HasVal(opIncRow(i)) And HasVal(daRow(i)) Then
            ebitdaRow(i) = opIncRow(i) + daRow(i)
        Else
            ebitdaRow(i) = ""
        End If

        If HasVal(evRow(i)) And HasVal(ebitdaRow(i)) And ebitdaRow(i) <> 0 Then
            evEbitdaRow(i) = evRow(i) / ebitdaRow(i)
        Else
            evEbitdaRow(i) = ""
        End If

        If HasVal(evRow(i)) And HasVal(revRow(i)) And revRow(i) <> 0 Then
            evSalesRow(i) = evRow(i) / revRow(i)
        Else
            evSalesRow(i) = ""
        End If

        If HasVal(netIncRow(i)) And HasVal(equityRow(i)) And equityRow(i) <> 0 Then
            roeRow(i) = netIncRow(i) / equityRow(i)
        Else
            roeRow(i) = ""
        End If

        ' NOPAT / invested capital; missing tax rate is approximated as 0
        ' (i.e. treated as pre-tax operating income) rather than blanking ROIC.
        ' Invested capital here still uses LT debt only (not +short-term) to
        ' match the metric's usual "permanent capital" convention.
        Dim taxForRoic As Double
        If HasVal(taxRateRow(i)) Then taxForRoic = taxRateRow(i) Else taxForRoic = 0

        ' Computed into a guarded variable BEFORE the If, not inline inside its
        ' condition: VBA's "+" throws Type Mismatch when adding a Variant
        ' holding "" to a number, and -- because VBA's "And" does not
        ' short-circuit -- that addition still runs even when an earlier
        ' HasVal(...) in the same condition has already evaluated to False.
        ' (Confirmed empirically: this is exactly what crashed on FTC Solar's
        ' FY2022 column, which has no LongTermDebt tag at all -- ltDebtRow(i)
        ' = "".)
        Dim investedCapital As Variant
        If HasVal(equityRow(i)) And HasVal(ltDebtRow(i)) Then
            investedCapital = equityRow(i) + ltDebtRow(i)
        Else
            investedCapital = ""
        End If

        If HasVal(opIncRow(i)) And HasVal(investedCapital) And investedCapital <> 0 Then
            roicRow(i) = (opIncRow(i) * (1 - taxForRoic)) / investedCapital
        Else
            roicRow(i) = ""
        End If

        If HasVal(dpsRow(i)) And HasVal(priceRow(i)) And priceRow(i) <> 0 Then
            divYieldRow(i) = dpsRow(i) / priceRow(i)
        Else
            divYieldRow(i) = ""
        End If

        If HasVal(dpsRow(i)) And HasVal(epsRow(i)) And epsRow(i) > 0 Then
            payoutRow(i) = dpsRow(i) / epsRow(i)
        Else
            payoutRow(i) = ""
        End If

        ' epsGrowthRow is stored as a decimal fraction (0.15 = 15%), so scale
        ' by 100 for the conventional PEG denominator. Blank/negative growth
        ' (incl. i=1, which has no prior year) leaves PEG blank rather than
        ' producing a misleading negative or divide-by-zero result.
        If HasVal(peRow(i)) And HasVal(epsGrowthRow(i)) And epsGrowthRow(i) > 0 Then
            pegRow(i) = peRow(i) / (epsGrowthRow(i) * 100)
        Else
            pegRow(i) = ""
        End If

        ' ---- DuPont ROE decomposition: ROE = Net Margin x Asset Turnover x Equity Multiplier ----
        If HasVal(netIncRow(i)) And HasVal(revRow(i)) And revRow(i) <> 0 Then
            netMarginRow(i) = netIncRow(i) / revRow(i)
        Else
            netMarginRow(i) = ""
        End If
        If HasVal(revRow(i)) And HasVal(assetsRow(i)) And assetsRow(i) <> 0 Then
            assetTurnoverRow(i) = revRow(i) / assetsRow(i)
        Else
            assetTurnoverRow(i) = ""
        End If
        If HasVal(assetsRow(i)) And HasVal(equityRow(i)) And equityRow(i) <> 0 Then
            equityMultiplierRow(i) = assetsRow(i) / equityRow(i)
        Else
            equityMultiplierRow(i) = ""
        End If

        ' ---- Cash Conversion Cycle: DIO + DSO - DPO, in days ----
        Dim daysInPeriod As Double
        If isQuarterly Then daysInPeriod = 91 Else daysInPeriod = 365

        If HasVal(invRow(i)) And HasVal(cogsRow(i)) And cogsRow(i) <> 0 Then
            dioRow(i) = invRow(i) / cogsRow(i) * daysInPeriod
        Else
            dioRow(i) = ""
        End If
        If HasVal(arRow(i)) And HasVal(revRow(i)) And revRow(i) <> 0 Then
            dsoRow(i) = arRow(i) / revRow(i) * daysInPeriod
        Else
            dsoRow(i) = ""
        End If
        If HasVal(apRow(i)) And HasVal(cogsRow(i)) And cogsRow(i) <> 0 Then
            dpoRow(i) = apRow(i) / cogsRow(i) * daysInPeriod
        Else
            dpoRow(i) = ""
        End If
        If HasVal(dioRow(i)) And HasVal(dsoRow(i)) And HasVal(dpoRow(i)) Then
            cccRow(i) = dioRow(i) + dsoRow(i) - dpoRow(i)
        Else
            cccRow(i) = ""
        End If

        ' ---- FCF conversion & yield ----
        If HasVal(fcfRow(i)) And HasVal(netIncRow(i)) And netIncRow(i) <> 0 Then
            fcfConversionRow(i) = fcfRow(i) / netIncRow(i)
        Else
            fcfConversionRow(i) = ""
        End If
        If HasVal(fcfRow(i)) And HasVal(mktCapRow(i)) And mktCapRow(i) <> 0 Then
            fcfYieldRow(i) = fcfRow(i) / mktCapRow(i)
        Else
            fcfYieldRow(i) = ""
        End If

        ' ---- Multi-year CAGR (Revenue/EPS/FCF, 3Y and 5Y) -- annual table only.
        ' For the quarterly table, i-3/i-5 would look back 3/5 QUARTERS, not
        ' years, which would silently mislabel the metric; same reasoning
        ' applies to the early annual columns of a still-young filer (e.g. only
        ' 1 10-K on file), which structurally cannot have 3/5 years of history
        ' yet regardless of ticker. Written as an explicit "N/A" (not a blank
        ' cell) so a reader can tell "not enough history for this metric" apart
        ' from "the fetch silently failed", which a truly empty cell would look
        ' identical to. (SafeCagr can still return "" on its own -- e.g. an
        ' endpoint that's zero/negative, where no CAGR is mathematically
        ' meaningful -- that case is left as a blank, not N/A, since it's a
        ' different situation: enough periods exist, just not a computable one.)
        If Not isQuarterly And i > 3 Then
            revCagr3Row(i) = SafeCagr(revRow(i - 3), revRow(i), 3)
            epsCagr3Row(i) = SafeCagr(epsRow(i - 3), epsRow(i), 3)
            fcfCagr3Row(i) = SafeCagr(fcfRow(i - 3), fcfRow(i), 3)
        Else
            revCagr3Row(i) = "N/A": epsCagr3Row(i) = "N/A": fcfCagr3Row(i) = "N/A"
        End If
        If Not isQuarterly And i > 5 Then
            revCagr5Row(i) = SafeCagr(revRow(i - 5), revRow(i), 5)
            epsCagr5Row(i) = SafeCagr(epsRow(i - 5), epsRow(i), 5)
            fcfCagr5Row(i) = SafeCagr(fcfRow(i - 5), fcfRow(i), 5)
        Else
            revCagr5Row(i) = "N/A": epsCagr5Row(i) = "N/A": fcfCagr5Row(i) = "N/A"
        End If
    Next i

    ' ---- Write the table, grouped into labeled sections instead of one flat
    ' list -- 34 rows in a single block was hard to scan once the valuation
    ' ratios were added. curRow is threaded through by reference so row numbers
    ' are never hardcoded (unlike the old fixed "row 22" style, which had to be
    ' hand-updated every time a row was added). ----
    Dim curRow As Long
    curRow = 2
    Dim bannerRows As Collection
    Set bannerRows = New Collection

    Call WriteSectionBanner(ws, curRow, "市場數據 Market Data", n + 1, bannerRows)
    Call WriteMetricRow(ws, curRow, "Stock Price", priceRow, n, "price")
    Call WriteMetricRow(ws, curRow, "Market Cap (M)", mktCapRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Common Shares (M)", sharesRow, n, "money")

    Call WriteSectionBanner(ws, curRow, "損益 Income Statement", n + 1, bannerRows)
    Call WriteMetricRow(ws, curRow, "Revenue (M)", revRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Sales Growth %", salesGrowthRow, n, "pct")
    Call WriteMetricRow(ws, curRow, "Revenue CAGR 3Y", revCagr3Row, n, "pct")
    Call WriteMetricRow(ws, curRow, "Revenue CAGR 5Y", revCagr5Row, n, "pct")
    Call WriteMetricRow(ws, curRow, "EPS (GAAP)", epsRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "EPS Growth %", epsGrowthRow, n, "pct")
    Call WriteMetricRow(ws, curRow, "EPS CAGR 3Y", epsCagr3Row, n, "pct")
    Call WriteMetricRow(ws, curRow, "EPS CAGR 5Y", epsCagr5Row, n, "pct")
    Call WriteMetricRow(ws, curRow, "Operating Income (M)", opIncRow, n, "money")
    Call WriteMetricRow(ws, curRow, "EBITDA (M)", ebitdaRow, n, "money")

    Call WriteSectionBanner(ws, curRow, "估值倍數 Valuation Multiples", n + 1, bannerRows)
    Call WriteMetricRow(ws, curRow, "P/E", peRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "P/S", psRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "Sales per Share", spsRow, n, "price")
    Call WriteMetricRow(ws, curRow, "EV/EBITDA", evEbitdaRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "EV/Sales", evSalesRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "PEG Ratio", pegRow, n, "ratio")

    Call WriteSectionBanner(ws, curRow, "獲利能力與資本回報 Profitability & Returns", n + 1, bannerRows)
    Call WriteMetricRow(ws, curRow, "ROE", roeRow, n, "pct")
    Call WriteMetricRow(ws, curRow, "  Net Margin (DuPont)", netMarginRow, n, "pct")
    Call WriteMetricRow(ws, curRow, "  Asset Turnover (DuPont)", assetTurnoverRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "  Equity Multiplier (DuPont)", equityMultiplierRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "ROIC", roicRow, n, "pct")
    Call WriteMetricRow(ws, curRow, "Effective Tax Rate", taxRateRow, n, "pct")

    Call WriteSectionBanner(ws, curRow, "營運資金 Working Capital", n + 1, bannerRows)
    Call WriteMetricRow(ws, curRow, "Inventory (M)", invRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Accounts Receivable (M)", arRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Accounts Payable (M)", apRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Current Assets (M)", curAssetsRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Current Liabilities (M)", curLiabRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Current Ratio", curRatioRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "Days Inventory Outstanding", dioRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "Days Sales Outstanding", dsoRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "Days Payable Outstanding", dpoRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "Cash Conversion Cycle (days)", cccRow, n, "ratio")

    Call WriteSectionBanner(ws, curRow, "資本結構 Capital Structure", n + 1, bannerRows)
    Call WriteMetricRow(ws, curRow, "Cash (M)", cashRow, n, "money")
    Call WriteMetricRow(ws, curRow, "LT Debt (M)", ltDebtRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Short-Term Debt (M)", stDebtRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Stockholders Equity (M)", equityRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Book Value/Share", bookValRow, n, "price")
    Call WriteMetricRow(ws, curRow, "Enterprise Value (M)", evRow, n, "money")

    Call WriteSectionBanner(ws, curRow, "現金流 Cash Flow", n + 1, bannerRows)
    Call WriteMetricRow(ws, curRow, "CapEx (M)", capexRow, n, "money")
    Call WriteMetricRow(ws, curRow, "D&A (M)", daRow, n, "money")
    Call WriteMetricRow(ws, curRow, "Free Cash Flow (M)", fcfRow, n, "money")
    Call WriteMetricRow(ws, curRow, "FCF Conversion (FCF/NI)", fcfConversionRow, n, "ratio")
    Call WriteMetricRow(ws, curRow, "FCF Yield", fcfYieldRow, n, "pct")
    Call WriteMetricRow(ws, curRow, "FCF CAGR 3Y", fcfCagr3Row, n, "pct")
    Call WriteMetricRow(ws, curRow, "FCF CAGR 5Y", fcfCagr5Row, n, "pct")

    Call WriteSectionBanner(ws, curRow, "股利 Dividends", n + 1, bannerRows)
    Call WriteMetricRow(ws, curRow, "Dividend Yield", divYieldRow, n, "pct")
    Call WriteMetricRow(ws, curRow, "Payout Ratio", payoutRow, n, "pct")

    Dim lastRow As Long
    lastRow = curRow - 1

    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, n + 1))
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Yu Gothic"
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n + 1)).Font.Color = RGB(255, 165, 0)
    Call ApplyGridBorder(ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, n + 1)))
    ' Section banners must be styled AFTER the blanket black-paint above, or
    ' this paint would immediately overwrite their gray background.
    Call StyleBannerRows(ws, bannerRows, n + 1)

    ' ---- P/E and P/S percentile band across all fetched fiscal years ----
    ' Not pinned to "5 years" -- uses whatever window was actually fetched
    ' (peRow/psRow already hold exactly the n fetched years), so it stays
    ' consistent with whatever the Input sheet's "10-K 年數" setting is.
    Dim bandCol As Long
    bandCol = n + 3
    ws.Cells(1, bandCol).Value = "分位數（近" & n & IIf(isQuarterly, "季", "年") & "）"
    ws.Cells(1, bandCol).Font.Bold = True
    ws.Cells(2, bandCol).Value = "P/E"
    ws.Cells(3, bandCol).Value = "P/S"
    ws.Cells(1, bandCol + 1).Value = "Min"
    ws.Cells(1, bandCol + 2).Value = "Median"
    ws.Cells(1, bandCol + 3).Value = "Max"

    Dim peFiltered As Variant, psFiltered As Variant
    peFiltered = FilteredArray(peRow, 1, n)
    psFiltered = FilteredArray(psRow, 1, n)

    If Not IsEmpty(peFiltered) Then
        ws.Cells(2, bandCol + 1).Value = Application.WorksheetFunction.Min(peFiltered)
        ws.Cells(2, bandCol + 2).Value = Application.WorksheetFunction.Median(peFiltered)
        ws.Cells(2, bandCol + 3).Value = Application.WorksheetFunction.Max(peFiltered)
    End If
    If Not IsEmpty(psFiltered) Then
        ws.Cells(3, bandCol + 1).Value = Application.WorksheetFunction.Min(psFiltered)
        ws.Cells(3, bandCol + 2).Value = Application.WorksheetFunction.Median(psFiltered)
        ws.Cells(3, bandCol + 3).Value = Application.WorksheetFunction.Max(psFiltered)
    End If
    ws.Range(ws.Cells(2, bandCol + 1), ws.Cells(3, bandCol + 3)).NumberFormat = "0.00"

    With ws.Range(ws.Cells(1, bandCol), ws.Cells(3, bandCol + 3))
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Yu Gothic"
    End With
    ws.Range(ws.Cells(1, bandCol), ws.Cells(1, bandCol + 3)).Font.Color = RGB(255, 165, 0)
    Call ApplyGridBorder(ws.Range(ws.Cells(1, bandCol), ws.Cells(3, bandCol + 3)))

    ' extend black background well beyond the actual table so scrolling/a wider
    ' window doesn't hit a white cutoff, matching every other sheet
    ws.Range("A1:AN500").Interior.Color = RGB(0, 0, 0)

    tableLastRow = lastRow
    If bandCol + 3 > n + 1 Then
        tableLastCol = bandCol + 3
    Else
        tableLastCol = n + 1
    End If
    Exit Sub

BSTIErr:
    ' Re-raise with the fiscal-year column and filing identifiers that were
    ' being processed when it failed -- "Type mismatch" alone gives no way to
    ' tell which of the n columns (or which concept lookup) is the culprit.
    Err.Raise Err.Number, "BuildSnapshotTableInto", _
        "column i=" & i & " accn=" & accn & " reportDate=" & reportDate & ": " & Err.Description
End Sub

' Returns a plain 1-based Double() array containing only the HasVal entries of
' src(fromIdx..toIdx), for feeding Application.WorksheetFunction.Min/Median/Max
' (which choke on a Variant array mixed with "" blanks). Returns Empty if none
' of the entries have a value.
Private Function FilteredArray(ByRef src As Variant, ByVal fromIdx As Long, ByVal toIdx As Long) As Variant
    Dim tmp() As Double
    ReDim tmp(1 To toIdx - fromIdx + 1)
    Dim cnt As Long
    cnt = 0
    Dim j As Long
    For j = fromIdx To toIdx
        If HasVal(src(j)) Then
            cnt = cnt + 1
            tmp(cnt) = CDbl(src(j))
        End If
    Next j
    If cnt = 0 Then
        FilteredArray = Empty
    Else
        Dim outArr() As Double
        ReDim outArr(1 To cnt)
        Dim k As Long
        For k = 1 To cnt
            outArr(k) = tmp(k)
        Next k
        FilteredArray = outArr
    End If
End Function

' ---------- Financial trend charts (one line chart per metric, annual & quarterly separated) ----------

' Reads every numeric metric column already written to the Filings sheet and
' draws a small line chart per metric, per period type (10-K rows = annual,
' 10-Q rows = quarterly). Filings sheet rows are 10-K block then 10-Q block,
' each in newest-first order (as fetched from SEC) -- charts read them in
' reverse so the trend runs oldest-to-newest, left-to-right, and that
' oldest-to-newest series is what gets written to RawData (starting at rawCol,
' each metric's annual block immediately followed by its quarterly one) for the
' chart to reference. startLeft/startTop set where the 2-column chart grid
' begins (so it can be stacked below other Dashboard content instead of always
' starting at the sheet's top-left).
' filingsLastRow is the 10-K/10-Q block's own last row on wsFilings, passed in
' explicitly rather than scanned via .End(xlUp) -- the Filings sheet now has
' 8-K/Form 4 rows (and a section banner) appended below that block, and
' scanning to the sheet's actual last used row would misread those as more
' quarterly financial data.
Private Sub BuildFinancialChartsInto(ByVal ws As Worksheet, ByVal wsFilings As Worksheet, ByVal filingsLastRow As Long, ByVal wsRaw As Worksheet, ByRef rawCol As Long, ByVal startLeft As Double, ByVal startTop As Double, Optional ByVal annualFormMarker As String = "10-K")
    Dim lastRow As Long
    lastRow = filingsLastRow
    If lastRow < 2 Then Exit Sub

    Dim kRows As Long
    kRows = 0
    Dim r As Long
    For r = 2 To lastRow
        If InStr(1, CStr(wsFilings.Cells(r, 3).Value), annualFormMarker, vbTextCompare) > 0 Then kRows = kRows + 1
    Next r

    Dim metricCols As Variant
    metricCols = Array(10, 11, 12, 13, 14, 15, 17, 18, 19, 20, 21, 22, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35)

    Dim leftPositions As Variant
    leftPositions = Array(startLeft, startLeft + 390)
    Dim posIdx As Long
    posIdx = 0
    Dim topPos As Double
    topPos = startTop
    Dim chartIdx As Long
    chartIdx = 0

    Dim mc As Variant
    For Each mc In metricCols
        Dim metricName As String
        metricName = CStr(wsFilings.Cells(1, CLng(mc)).Value)
        If metricName <> "" Then
            If kRows > 0 Then
                Call BuildOneMetricChart(ws, wsFilings, wsRaw, rawCol, 2, 1 + kRows, CLng(mc), metricName & "（年度 Annual）", CDbl(leftPositions(posIdx)), topPos, chartIdx)
                chartIdx = chartIdx + 1
                posIdx = (posIdx + 1) Mod 2
                If posIdx = 0 Then topPos = topPos + 230
            End If
            If lastRow >= 2 + kRows Then
                Call BuildOneMetricChart(ws, wsFilings, wsRaw, rawCol, 2 + kRows, lastRow, CLng(mc), metricName & "（季度 Quarterly）", CDbl(leftPositions(posIdx)), topPos, chartIdx)
                chartIdx = chartIdx + 1
                posIdx = (posIdx + 1) Mod 2
                If posIdx = 0 Then topPos = topPos + 230
            End If
        End If
    Next mc

    ' Background paint sized to this grid's own pixel depth (default row height
    ' ~15pt). This is now the only paint that reaches below the snapshot table's
    ' own A1:AN500 -- the price-history data table that used to paint this
    ' region on its way past has moved to the RawData sheet. 60 columns wide to
    ' match what that table's ApplyDarkTheme call used to cover.
    Dim estRows As Long
    estRows = CLng(topPos / 15) + 200
    If estRows < 500 Then estRows = 500
    ws.Range(ws.Cells(1, 1), ws.Cells(estRows, 60)).Interior.Color = RGB(0, 0, 0)
End Sub

Private Sub BuildOneMetricChart(ByVal ws As Worksheet, ByVal wsFilings As Worksheet, ByVal wsRaw As Worksheet, ByRef rawCol As Long, ByVal firstRow As Long, ByVal lastRow As Long, ByVal metricCol As Long, ByVal chartTitle As String, ByVal leftPos As Double, ByVal topPos As Double, ByVal paletteIdx As Long)
    Dim n As Long
    n = lastRow - firstRow + 1
    If n <= 0 Then Exit Sub

    Dim vals() As Variant
    Dim labels() As Variant
    ReDim vals(1 To n)
    ReDim labels(1 To n)

    Dim r As Long, idx As Long
    idx = 0
    For r = lastRow To firstRow Step -1
        idx = idx + 1
        Dim v As Variant
        v = wsFilings.Cells(r, metricCol).Value
        If IsNumeric(v) Then vals(idx) = CDbl(v) Else vals(idx) = 0

        Dim fy As String, fp As String
        fy = CStr(wsFilings.Cells(r, 4).Value)
        fp = CStr(wsFilings.Cells(r, 5).Value)
        If fp = "FY" Or fp = "" Then
            labels(idx) = fy
        Else
            labels(idx) = fy & fp
        End If
    Next r

    ' Carry the Filings column's own number format across so a margin stays a
    ' percentage and a revenue stays a thousands-separated integer on RawData
    ' too, instead of every metric landing as a bare General number.
    Dim keyRng As Range, valRng As Range
    Call WriteRawBlock(wsRaw, rawCol, chartTitle, "期間 Period", CStr(wsFilings.Cells(1, metricCol).Value), _
        labels, vals, n, "@", CStr(wsFilings.Cells(firstRow, metricCol).NumberFormat), keyRng, valRng)
    If valRng Is Nothing Then Exit Sub

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=leftPos, Top:=topPos, Width:=380, Height:=210)
    Dim ch As Chart
    Set ch = co.Chart
    Call ClearAutoSeries(ch)
    ch.ChartType = xlLine

    Dim ser As Series
    Set ser = ch.SeriesCollection.NewSeries
    ser.Values = valRng
    ser.XValues = keyRng
    ser.Name = chartTitle

    ch.HasTitle = True
    ch.ChartTitle.Text = chartTitle
    ch.HasLegend = False

    Call ApplyDarkThemeToChart(ch, Array(OrangePalette(paletteIdx)))

    ' Same fit as the price chart: these are trend lines, and Excel's zero-based
    ' auto-scale flattens most of them (股東權益 179M-275M on a 0-300M axis).
    ' A non-zero baseline is the right call for a trend line specifically -- it
    ' would be misleading on a bar chart, which none of these are.
    Dim axLo As Double, axHi As Double
    If ArrayMinMax(vals, n, axLo, axHi) Then Call FitValueAxis(ch.Axes(xlValue), axLo, axHi)
End Sub
