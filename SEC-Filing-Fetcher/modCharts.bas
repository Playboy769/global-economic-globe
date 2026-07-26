Attribute VB_Name = "modCharts"
Option Explicit

' "No value" is represented as a zero-length string sentinel inside Variant
' arrays here. LookupConceptValue always returns its result typed as a String
' subtype Variant (even numeric results), so checking VarType against vbString
' can't distinguish "has a value" from "empty" -- go through CStr()/IsNumeric()
' instead, which also sidesteps the Type Mismatch a direct "v = ''" comparison
' would throw when v holds a Double.
Private Function HasVal(ByVal v As Variant) As Boolean
    HasVal = (Not IsNull(v)) And (CStr(v) <> "") And IsNumeric(v)
End Function

' Finds or creates the sheet but never clears it -- callers that need a clean
' slate must call ClearSheetForRebuild themselves, once, before writing. This
' split exists so multiple builders can safely share one sheet (Dashboard):
' each GetOrCreateSheet call used to wipe the whole sheet unconditionally,
' which would destroy an earlier builder's content on a shared sheet.
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
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

' ---------- Dashboard orchestrator ----------

' Owns the single "Dashboard" sheet: clears it once, then stacks the snapshot
' table, the price/EPS chart, and the financial trend chart grid vertically
' (table width varies with how many fiscal years were fetched, so vertical
' stacking sidesteps needing a dynamic side-by-side width calculation).
Public Sub BuildDashboard(ByVal wb As Workbook, ByVal ticker As String, ByVal entityName As String, _
    ByVal filings10K As Collection, ByVal filings10Q As Collection, ByVal mapEps As Object, ByVal prices As Object, _
    ByVal mapRevenue As Object, ByVal mapShares As Object, ByVal allMaps As Variant, _
    ByVal mapInventory As Object, ByVal mapAR As Object, ByVal mapCurrentAssets As Object, ByVal mapCurrentLiabilities As Object, _
    ByVal mapLongTermDebt As Object, ByVal mapStockholdersEquity As Object, ByVal mapEffectiveTaxRate As Object, ByVal mapCapEx As Object, ByVal mapCFO As Object, _
    ByVal mapCash As Object, ByVal mapDA As Object, ByVal mapOperatingIncome As Object, ByVal mapDividends As Object, ByVal mapNetIncome As Object, _
    ByVal wsFilings As Worksheet)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, "Dashboard")
    Call ClearSheetForRebuild(ws)

    Dim tableLastRow As Long, tableLastCol As Long
    Call BuildSnapshotTableInto(ws, entityName, ticker, filings10K, mapRevenue, mapEps, mapShares, allMaps, prices, _
        mapInventory, mapAR, mapCurrentAssets, mapCurrentLiabilities, mapLongTermDebt, mapStockholdersEquity, mapEffectiveTaxRate, mapCapEx, mapCFO, _
        mapCash, mapDA, mapOperatingIncome, mapDividends, mapNetIncome, tableLastRow, tableLastCol)

    Dim priceChartTop As Double, priceChartBottom As Double
    priceChartTop = ws.Cells(tableLastRow + 3, 1).Top
    Dim priceChartBottomOut As Double
    Call BuildPriceChartInto(ws, ticker, filings10K, filings10Q, mapEps, prices, tableLastRow + 3, 300, priceChartTop, priceChartBottomOut)
    priceChartBottom = priceChartBottomOut

    Call BuildFinancialChartsInto(ws, wsFilings, 10, priceChartBottom + 20)

    ws.Columns.AutoFit
End Sub

' Chart 1: last ~5 months of daily closing price (Stooq) plotted against each
' filing's reported diluted EPS, held flat until the next filing supersedes it
' (a step function), so the reader can see price movement against what was
' actually known/reported at each point in time. No Forward P/E series.
' startRow: the row the small date/EPS/price data table (chart's data source)
' begins at. chartLeft/chartTop: pixel position of the floating chart object
' itself (independent of the data table's row/column position). chartBottomOut
' reports the chart's pixel bottom edge so the caller can stack the next block
' below it.
Private Sub BuildPriceChartInto(ByVal ws As Worksheet, ByVal ticker As String, ByVal filings10K As Collection, ByVal filings10Q As Collection, _
    ByVal mapEps As Object, ByVal prices As Object, ByVal startRow As Long, ByVal chartLeft As Double, ByVal chartTop As Double, ByRef chartBottomOut As Double)

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
        ws.Cells(startRow, 1).Value = "找不到 " & ticker & " 的股價資料（Stooq 無資料，可能不是美股代號）"
        Exit Sub
    End If

    ws.Cells(startRow, 1).Value = "日期"
    ws.Cells(startRow, 2).Value = "Diluted EPS (as reported)"
    ws.Cells(startRow, 3).Value = ticker & " 股價"
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 3)).Font.Bold = True

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

    Dim rowOut As Long
    rowOut = startRow + 1
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

            ws.Cells(rowOut, 1).Value = CDate(curDateStr)
            If HasVal(stepEps) Then ws.Cells(rowOut, 2).Value = stepEps
            ws.Cells(rowOut, 3).Value = prices(curDateStr)
            rowOut = rowOut + 1
        End If
    Next kk

    If rowOut = startRow + 1 Then
        ws.Cells(startRow + 1, 1).Value = "近5年無股價資料"
        Exit Sub
    End If

    ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(rowOut - 1, 1)).NumberFormat = "yyyy-mm-dd"
    Call ApplyDarkTheme(ws, ws.Range(ws.Cells(startRow, 1), ws.Cells(rowOut - 1, 3)), startRow)

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=750, Height:=400)
    Dim ch As Chart
    Set ch = co.Chart
    ch.SetSourceData Source:=ws.Range(ws.Cells(startRow, 1), ws.Cells(rowOut - 1, 3))
    ch.ChartType = xlLine

    ch.SeriesCollection(1).ChartType = xlArea
    ch.SeriesCollection(1).Name = "Diluted EPS (as reported)"

    ch.SeriesCollection(2).ChartType = xlLine
    ch.SeriesCollection(2).Name = ticker & " Price"
    ch.SeriesCollection(2).AxisGroup = xlSecondary
    ch.SeriesCollection(2).Format.Line.Weight = 0.5

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

    chartBottomOut = co.Top + co.Height
End Sub

' Table: actual historical fiscal-year snapshot (Stock Price / Market Cap /
' Revenue / EPS / P-E / P-S / growth rates / valuation & return ratios) across
' the 10-Ks already fetched, plus a trailing P/E and P/S percentile band.
' Forecast years are intentionally excluded -- SEC filings only report what has
' actually happened. tableLastRow/tableLastCol report the bottom-right cell
' actually used so the caller can stack the next block below it.
Private Sub BuildSnapshotTableInto(ByVal ws As Worksheet, ByVal entityName As String, ByVal ticker As String, ByVal filings10K As Collection, _
    ByVal mapRevenue As Object, ByVal mapEps As Object, ByVal mapShares As Object, ByVal allMaps As Variant, ByVal prices As Object, _
    ByVal mapInventory As Object, ByVal mapAR As Object, ByVal mapCurrentAssets As Object, ByVal mapCurrentLiabilities As Object, _
    ByVal mapLongTermDebt As Object, ByVal mapStockholdersEquity As Object, ByVal mapEffectiveTaxRate As Object, ByVal mapCapEx As Object, ByVal mapCFO As Object, _
    ByVal mapCash As Object, ByVal mapDA As Object, ByVal mapOperatingIncome As Object, ByVal mapDividends As Object, ByVal mapNetIncome As Object, _
    ByRef tableLastRow As Long, ByRef tableLastCol As Long)

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
    Dim i As Long
    i = 0
    Dim f As Variant
    For Each f In filings10K
        i = i + 1
        Set ordered(n - i + 1) = f
    Next f

    ws.Range("A1").Value = entityName & " (" & ticker & ")"
    ws.Range("A1").Font.Bold = True

    Dim labels As Variant
    labels = Array("", "Stock Price", "Market Cap", "Common Shares", "Revenue (M)", "Sales Growth %", "EPS (GAAP)", "EPS Growth %", "P/E", "P/S", "Sales per Share", _
        "Inventory (M)", "Accounts Receivable (M)", "Current Assets (M)", "Current Liabilities (M)", "Current Ratio", _
        "LT Debt (M)", "Stockholders Equity (M)", "Book Value/Share", "Effective Tax Rate", "CapEx (M)", "Free Cash Flow (M)", _
        "Cash (M)", "Operating Income (M)", "D&A (M)", "Enterprise Value (M)", "EBITDA (M)", "EV/EBITDA", "EV/Sales", _
        "ROE", "ROIC", "Dividend Yield", "Payout Ratio", "PEG Ratio")
    For i = 1 To UBound(labels)
        ws.Cells(i + 1, 1).Value = labels(i)
    Next i

    Dim priceRow() As Variant, mktCapRow() As Variant, sharesRow() As Variant, revRow() As Variant
    Dim epsRow() As Variant, salesGrowthRow() As Variant, epsGrowthRow() As Variant
    Dim peRow() As Variant, psRow() As Variant, spsRow() As Variant
    Dim invRow() As Variant, arRow() As Variant, curAssetsRow() As Variant, curLiabRow() As Variant, curRatioRow() As Variant
    Dim ltDebtRow() As Variant, equityRow() As Variant, bookValRow() As Variant, taxRateRow() As Variant, capexRow() As Variant, fcfRow() As Variant
    Dim cashRow() As Variant, opIncRow() As Variant, daRow() As Variant, evRow() As Variant, ebitdaRow() As Variant
    Dim evEbitdaRow() As Variant, evSalesRow() As Variant, roeRow() As Variant, roicRow() As Variant
    Dim dpsRow() As Variant, netIncRow() As Variant, divYieldRow() As Variant, payoutRow() As Variant, pegRow() As Variant
    ReDim priceRow(1 To n): ReDim mktCapRow(1 To n): ReDim sharesRow(1 To n): ReDim revRow(1 To n)
    ReDim epsRow(1 To n): ReDim salesGrowthRow(1 To n): ReDim epsGrowthRow(1 To n)
    ReDim peRow(1 To n): ReDim psRow(1 To n): ReDim spsRow(1 To n)
    ReDim invRow(1 To n): ReDim arRow(1 To n): ReDim curAssetsRow(1 To n): ReDim curLiabRow(1 To n): ReDim curRatioRow(1 To n)
    ReDim ltDebtRow(1 To n): ReDim equityRow(1 To n): ReDim bookValRow(1 To n): ReDim taxRateRow(1 To n): ReDim capexRow(1 To n): ReDim fcfRow(1 To n)
    ReDim cashRow(1 To n): ReDim opIncRow(1 To n): ReDim daRow(1 To n): ReDim evRow(1 To n): ReDim ebitdaRow(1 To n)
    ReDim evEbitdaRow(1 To n): ReDim evSalesRow(1 To n): ReDim roeRow(1 To n): ReDim roicRow(1 To n)
    ReDim dpsRow(1 To n): ReDim netIncRow(1 To n): ReDim divYieldRow(1 To n): ReDim payoutRow(1 To n): ReDim pegRow(1 To n)

    For i = 1 To n
        Dim accn As String, reportDate As String, form As String
        accn = ordered(i)("accn")
        reportDate = ordered(i)("reportDate")
        form = ordered(i)("form")

        Dim fyfp As String, parts As Variant, colHeader As String
        fyfp = LookupFyFp(allMaps, accn, reportDate, form)
        parts = Split(fyfp, "|")
        If parts(0) <> "" Then
            colHeader = "FY" & parts(0)
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
        Dim ltDebtV As Variant, equityV As Variant, taxRateV As Variant, capexV As Variant, cfoV As Variant
        invV = LookupConceptValue(mapInventory, accn, reportDate, form)
        arV = LookupConceptValue(mapAR, accn, reportDate, form)
        curAssetsV = LookupConceptValue(mapCurrentAssets, accn, reportDate, form)
        curLiabV = LookupConceptValue(mapCurrentLiabilities, accn, reportDate, form)
        ltDebtV = LookupConceptValue(mapLongTermDebt, accn, reportDate, form)
        equityV = LookupConceptValue(mapStockholdersEquity, accn, reportDate, form)
        taxRateV = LookupConceptValue(mapEffectiveTaxRate, accn, reportDate, form)
        capexV = LookupConceptValue(mapCapEx, accn, reportDate, form)
        cfoV = LookupConceptValue(mapCFO, accn, reportDate, form)

        If HasVal(invV) Then invRow(i) = CDbl(invV) Else invRow(i) = ""
        If HasVal(arV) Then arRow(i) = CDbl(arV) Else arRow(i) = ""
        If HasVal(curAssetsV) Then curAssetsRow(i) = CDbl(curAssetsV) Else curAssetsRow(i) = ""
        If HasVal(curLiabV) Then curLiabRow(i) = CDbl(curLiabV) Else curLiabRow(i) = ""
        If HasVal(ltDebtV) Then ltDebtRow(i) = CDbl(ltDebtV) Else ltDebtRow(i) = ""
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

        ' ---- Valuation & return ratios (new) ----
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

        ' Approximation: no short-term/current-debt concept is fetched anywhere
        ' in this workbook (only LongTermDebt(Noncurrent)), so EV omits the
        ' current portion of debt and will run slightly low.
        If HasVal(mktCapRow(i)) And HasVal(ltDebtRow(i)) And HasVal(cashRow(i)) Then
            evRow(i) = mktCapRow(i) + ltDebtRow(i) - cashRow(i)
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
        Dim taxForRoic As Double
        If HasVal(taxRateRow(i)) Then taxForRoic = taxRateRow(i) Else taxForRoic = 0
        If HasVal(opIncRow(i)) And HasVal(equityRow(i)) And HasVal(ltDebtRow(i)) And (equityRow(i) + ltDebtRow(i)) <> 0 Then
            roicRow(i) = (opIncRow(i) * (1 - taxForRoic)) / (equityRow(i) + ltDebtRow(i))
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

        ws.Cells(2, i + 1).Value = priceRow(i)
        ws.Cells(3, i + 1).Value = mktCapRow(i)
        ws.Cells(4, i + 1).Value = sharesRow(i)
        ws.Cells(5, i + 1).Value = revRow(i)
        ws.Cells(6, i + 1).Value = salesGrowthRow(i)
        ws.Cells(7, i + 1).Value = epsRow(i)
        ws.Cells(8, i + 1).Value = epsGrowthRow(i)
        ws.Cells(9, i + 1).Value = peRow(i)
        ws.Cells(10, i + 1).Value = psRow(i)
        ws.Cells(11, i + 1).Value = spsRow(i)
        ws.Cells(12, i + 1).Value = invRow(i)
        ws.Cells(13, i + 1).Value = arRow(i)
        ws.Cells(14, i + 1).Value = curAssetsRow(i)
        ws.Cells(15, i + 1).Value = curLiabRow(i)
        ws.Cells(16, i + 1).Value = curRatioRow(i)
        ws.Cells(17, i + 1).Value = ltDebtRow(i)
        ws.Cells(18, i + 1).Value = equityRow(i)
        ws.Cells(19, i + 1).Value = bookValRow(i)
        ws.Cells(20, i + 1).Value = taxRateRow(i)
        ws.Cells(21, i + 1).Value = capexRow(i)
        ws.Cells(22, i + 1).Value = fcfRow(i)
        ws.Cells(23, i + 1).Value = cashRow(i)
        ws.Cells(24, i + 1).Value = opIncRow(i)
        ws.Cells(25, i + 1).Value = daRow(i)
        ws.Cells(26, i + 1).Value = evRow(i)
        ws.Cells(27, i + 1).Value = ebitdaRow(i)
        ws.Cells(28, i + 1).Value = evEbitdaRow(i)
        ws.Cells(29, i + 1).Value = evSalesRow(i)
        ws.Cells(30, i + 1).Value = roeRow(i)
        ws.Cells(31, i + 1).Value = roicRow(i)
        ws.Cells(32, i + 1).Value = divYieldRow(i)
        ws.Cells(33, i + 1).Value = payoutRow(i)
        ws.Cells(34, i + 1).Value = pegRow(i)
    Next i

    ws.Range(ws.Cells(2, 2), ws.Cells(2, n + 1)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(3, 2), ws.Cells(3, n + 1)).NumberFormat = "#,##0,,""M"""
    ws.Range(ws.Cells(4, 2), ws.Cells(4, n + 1)).NumberFormat = "#,##0,,""M"""
    ws.Range(ws.Cells(5, 2), ws.Cells(5, n + 1)).NumberFormat = "#,##0,,""M"""
    With ws.Range(ws.Cells(6, 2), ws.Cells(6, n + 1))
        .NumberFormat = "0.00%"
        .Font.Italic = True
    End With
    ws.Range(ws.Cells(7, 2), ws.Cells(7, n + 1)).NumberFormat = "0.00"
    With ws.Range(ws.Cells(8, 2), ws.Cells(8, n + 1))
        .NumberFormat = "0.00%"
        .Font.Italic = True
    End With
    ws.Range(ws.Cells(9, 2), ws.Cells(9, n + 1)).NumberFormat = "0.00"
    ws.Range(ws.Cells(10, 2), ws.Cells(10, n + 1)).NumberFormat = "0.00"
    ws.Range(ws.Cells(11, 2), ws.Cells(11, n + 1)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(12, 2), ws.Cells(15, n + 1)).NumberFormat = "#,##0,,""M"""
    ws.Range(ws.Cells(16, 2), ws.Cells(16, n + 1)).NumberFormat = "0.00"
    ws.Range(ws.Cells(17, 2), ws.Cells(18, n + 1)).NumberFormat = "#,##0,,""M"""
    ws.Range(ws.Cells(19, 2), ws.Cells(19, n + 1)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(20, 2), ws.Cells(20, n + 1)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(21, 2), ws.Cells(22, n + 1)).NumberFormat = "#,##0,,""M"""
    ws.Range(ws.Cells(23, 2), ws.Cells(27, n + 1)).NumberFormat = "#,##0,,""M"""
    ws.Range(ws.Cells(28, 2), ws.Cells(29, n + 1)).NumberFormat = "0.00"
    With ws.Range(ws.Cells(30, 2), ws.Cells(30, n + 1))
        .NumberFormat = "0.00%"
        .Font.Italic = True
    End With
    With ws.Range(ws.Cells(31, 2), ws.Cells(31, n + 1))
        .NumberFormat = "0.00%"
        .Font.Italic = True
    End With
    With ws.Range(ws.Cells(32, 2), ws.Cells(32, n + 1))
        .NumberFormat = "0.00%"
        .Font.Italic = True
    End With
    With ws.Range(ws.Cells(33, 2), ws.Cells(33, n + 1))
        .NumberFormat = "0.00%"
        .Font.Italic = True
    End With
    ws.Range(ws.Cells(34, 2), ws.Cells(34, n + 1)).NumberFormat = "0.00"

    With ws.Range(ws.Cells(1, 1), ws.Cells(34, n + 1))
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Yu Gothic"
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n + 1)).Font.Color = RGB(255, 165, 0)
    Call ApplyGridBorder(ws.Range(ws.Cells(1, 1), ws.Cells(34, n + 1)))

    ' ---- P/E and P/S percentile band across all fetched fiscal years ----
    ' Not pinned to "5 years" -- uses whatever window was actually fetched
    ' (peRow/psRow already hold exactly the n fetched years), so it stays
    ' consistent with whatever the Input sheet's "10-K 年數" setting is.
    Dim bandCol As Long
    bandCol = n + 3
    ws.Cells(1, bandCol).Value = "分位數（近" & n & "年）"
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

    tableLastRow = 34
    If bandCol + 3 > n + 1 Then
        tableLastCol = bandCol + 3
    Else
        tableLastCol = n + 1
    End If
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
' reverse so the trend runs oldest-to-newest, left-to-right. startLeft/startTop
' set where the 2-column chart grid begins (so it can be stacked below other
' Dashboard content instead of always starting at the sheet's top-left).
Private Sub BuildFinancialChartsInto(ByVal ws As Worksheet, ByVal wsFilings As Worksheet, ByVal startLeft As Double, ByVal startTop As Double)
    Dim lastRow As Long
    lastRow = wsFilings.Cells(wsFilings.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim kRows As Long
    kRows = 0
    Dim r As Long
    For r = 2 To lastRow
        If InStr(1, CStr(wsFilings.Cells(r, 3).Value), "10-K", vbTextCompare) > 0 Then kRows = kRows + 1
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
                Call BuildOneMetricChart(ws, wsFilings, 2, 1 + kRows, CLng(mc), metricName & "（年度 Annual）", CDbl(leftPositions(posIdx)), topPos, chartIdx)
                chartIdx = chartIdx + 1
                posIdx = (posIdx + 1) Mod 2
                If posIdx = 0 Then topPos = topPos + 230
            End If
            If lastRow >= 2 + kRows Then
                Call BuildOneMetricChart(ws, wsFilings, 2 + kRows, lastRow, CLng(mc), metricName & "（季度 Quarterly）", CDbl(leftPositions(posIdx)), topPos, chartIdx)
                chartIdx = chartIdx + 1
                posIdx = (posIdx + 1) Mod 2
                If posIdx = 0 Then topPos = topPos + 230
            End If
        End If
    Next mc

    ' Safety-net background paint sized to this grid's own pixel depth (default
    ' row height ~15pt), in case the price chart above it didn't already reach
    ' this far down (e.g. a ticker with no Stooq price data skips that paint).
    Dim estRows As Long
    estRows = CLng(topPos / 15) + 200
    If estRows < 500 Then estRows = 500
    ws.Range(ws.Cells(1, 1), ws.Cells(estRows, 40)).Interior.Color = RGB(0, 0, 0)
End Sub

Private Sub BuildOneMetricChart(ByVal ws As Worksheet, ByVal wsFilings As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal metricCol As Long, ByVal chartTitle As String, ByVal leftPos As Double, ByVal topPos As Double, ByVal paletteIdx As Long)
    Dim n As Long
    n = lastRow - firstRow + 1
    If n <= 0 Then Exit Sub

    Dim vals() As Double
    Dim labels() As String
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

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=leftPos, Top:=topPos, Width:=380, Height:=210)
    Dim ch As Chart
    Set ch = co.Chart
    ch.ChartType = xlLine

    Dim ser As Series
    Set ser = ch.SeriesCollection.NewSeries
    ser.Values = vals
    ser.XValues = labels
    ser.Name = chartTitle

    ch.HasTitle = True
    ch.ChartTitle.Text = chartTitle
    ch.HasLegend = False

    Call ApplyDarkThemeToChart(ch, Array(OrangePalette(paletteIdx)))
End Sub
