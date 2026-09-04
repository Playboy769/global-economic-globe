Attribute VB_Name = "modValuation"
Option Explicit

' ---------------------------------------------------------------------------
' Valuation add-ons shared by the US (modSEC) and TW (modMOPS) pipelines:
'   * customer advances / deposits (B/S 預收貨款)
'   * sustainable growth rate      (ROE x retention ratio, 內生成長率)
'   * ROIC vs WACC                 (NOPAT / invested capital, CAPM-derived WACC)
'
' Only WACC needs something the filings themselves don't carry: SEC and MOPS
' never disclose a cost of capital, so it is estimated here via CAPM. Beta is
' regressed from actual monthly price returns against a benchmark index; Rf,
' ERP and the tax fallback come from the Input sheet so they stay visible and
' adjustable instead of being buried as constants (see LoadCapmParams).
'
' Definitions chosen (documented here because several defensible ones exist):
'   Invested capital = Total assets - Current liabilities   ("capital employed")
'   NOPAT            = Operating income x (1 - effective tax rate)
'   ROIC             = NOPAT / invested capital
'   Retention ratio  = 1 - (dividend per share / diluted EPS)
'   Sustainable g    = ROE(period-end equity) x retention ratio
'
' Quarterly rows ANNUALISE the two flow-driven rates (ROIC, sustainable growth)
' by x4 so they sit on the same annual scale as WACC -- comparing an unscaled
' quarterly ROIC against an annual WACC would make almost every company look
' value-destroying. Balance-sheet columns (advances, invested capital) are
' point-in-time and are never scaled.
' ---------------------------------------------------------------------------

' Input-sheet cells holding the CAPM inputs. Kept in one place so build.ps1's
' sheet layout and the readers below can't drift apart.
Public Const CAPM_CELL_RF_US As String = "B12"
Public Const CAPM_CELL_ERP_US As String = "B13"
Public Const CAPM_CELL_TAX_US As String = "B14"
Public Const CAPM_CELL_INDEX_US As String = "B15"
Public Const CAPM_CELL_RF_TW As String = "B16"
Public Const CAPM_CELL_ERP_TW As String = "B17"
Public Const CAPM_CELL_TAX_TW As String = "B18"
Public Const CAPM_CELL_INDEX_TW As String = "B19"
Public Const CAPM_CELL_BETA_MONTHS As String = "B20"

' Column indices appended to the Filings / TW_Filings sheets. Both sheets share
' one layout (see modSEC.WriteHeaders / modMOPS.WriteTWHeaders), and columns
' 1-48 were already spoken for -- including 48 = Items, which several call
' sites address by literal number -- so these are appended past it rather than
' inserted, which would have silently shifted Items out from under those.
Public Const COL_ADV_CURRENT As Long = 49
Public Const COL_ADV_NONCURRENT As Long = 50
Public Const COL_ADV_TOTAL As Long = 51
Public Const COL_RETENTION As Long = 52
Public Const COL_SUSTAINABLE_G As Long = 53
Public Const COL_INVESTED_CAPITAL As Long = 54
Public Const COL_NOPAT As Long = 55
Public Const COL_ROIC As Long = 56
Public Const COL_WACC As Long = 57
Public Const COL_SPREAD As Long = 58
Public Const COL_ADV_SOURCE As Long = 59
Public Const COL_VALUATION_LAST As Long = 59

' Source labels written into COL_ADV_SOURCE, so a reader can tell at a glance
' whether a row's prepayment figure is the narrow customer-advance concept or
' the broader contract-liability one -- the two are NOT comparable across rows.
Public Const ADV_SRC_ADVANCES As String = "預收貨款 Customer Advances"
Public Const ADV_SRC_CONTRACT As String = "合約負債 Contract Liabilities"

' Header labels for columns 49-58, in order. Shared by both sheets.
Public Function ValuationHeaders() As Variant
    ValuationHeaders = Array( _
        "預收貨款-流動 Customer Advances (Current)", _
        "預收貨款-非流動 Customer Advances (Noncurrent)", _
        "預收貨款合計 Customer Advances Total", _
        "盈餘保留率 Retention Ratio", _
        "內生成長率 Sustainable Growth (年化)", _
        "投入資本 Invested Capital", _
        "稅後營業利益 NOPAT (年化)", _
        "投入資本報酬率 ROIC (年化)", _
        "加權平均資金成本 WACC", _
        "ROIC-WACC 價差 Spread", _
        "預收款來源 Prepayment Source")
End Function

' Picks the prepayment figures for one row: the narrow customer-advance tags
' when the filer uses them, otherwise the contract-liability ones.
'
' The fallback exists because the narrow concept is nearly unused: at CY2024Q4
' only ~43 SEC filers tagged CustomerAdvancesCurrent or CustomerDepositsCurrent
' (SEC frames API), against ~2,033 for ContractWithCustomerLiabilityCurrent,
' and the TW IFRS taxonomy has no customer-advance balance-sheet element at
' all. Advances-only would therefore have left these columns blank for the vast
' majority of companies in both markets. The fallback is gap-fill only, never
' an override, and every row records which concept it actually used.
Public Function PickAdvances(ByVal pCur As Variant, ByVal pNon As Variant, _
    ByVal fCur As Variant, ByVal fNon As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim havePrimary As Boolean
    havePrimary = IsNumeric(pCur)
    If Not havePrimary Then havePrimary = IsNumeric(pNon)

    If havePrimary Then
        d("cur") = pCur
        d("non") = pNon
        d("source") = ADV_SRC_ADVANCES
    Else
        Dim haveFallback As Boolean
        haveFallback = IsNumeric(fCur)
        If Not haveFallback Then haveFallback = IsNumeric(fNon)
        d("cur") = fCur
        d("non") = fNon
        If haveFallback Then
            d("source") = ADV_SRC_CONTRACT
        Else
            d("source") = ""
        End If
    End If

    Set PickAdvances = d
End Function

' ---------- CAPM parameter block ----------

' Reads the Input sheet's CAPM cells into one dictionary. Blank cells fall back
' to the defaults below rather than erroring, matching modSEC.GetSetting's
' "leave it blank and you get the default" convention.
Public Function LoadCapmParams(ByVal wsIn As Worksheet, ByVal isTaiwan As Boolean) As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")

    If isTaiwan Then
        p("rf") = ReadRate(wsIn, CAPM_CELL_RF_TW, 0.015)
        p("erp") = ReadRate(wsIn, CAPM_CELL_ERP_TW, 0.06)
        p("taxFallback") = ReadRate(wsIn, CAPM_CELL_TAX_TW, 0.2)
        p("indexSymbol") = ReadText(wsIn, CAPM_CELL_INDEX_TW, "^TWII")
    Else
        p("rf") = ReadRate(wsIn, CAPM_CELL_RF_US, 0.042)
        p("erp") = ReadRate(wsIn, CAPM_CELL_ERP_US, 0.05)
        p("taxFallback") = ReadRate(wsIn, CAPM_CELL_TAX_US, 0.21)
        p("indexSymbol") = ReadText(wsIn, CAPM_CELL_INDEX_US, "^GSPC")
    End If

    Dim months As Double
    months = ReadRate(wsIn, CAPM_CELL_BETA_MONTHS, 60)
    If months < 24 Then months = 24
    If months > 180 Then months = 180
    p("betaMonths") = CLng(months)

    ' Filled in later by AttachBeta once a ticker is known. 1.0 is the
    ' documented fallback (market beta) for a ticker with too little price
    ' history to regress -- never a silent zero, which would make Ke = Rf and
    ' quietly turn every ROIC-WACC spread positive.
    p("beta") = 1#
    p("betaIsEstimated") = False

    Set LoadCapmParams = p
End Function

Private Function ReadRate(ByVal wsIn As Worksheet, ByVal cellAddr As String, ByVal defaultVal As Double) As Double
    Dim v As Variant
    v = wsIn.Range(cellAddr).Value
    If IsEmpty(v) Then
        ReadRate = defaultVal
        Exit Function
    End If
    If Not IsNumeric(v) Then
        ReadRate = defaultVal
        Exit Function
    End If
    ReadRate = CDbl(v)
End Function

Private Function ReadText(ByVal wsIn As Worksheet, ByVal cellAddr As String, ByVal defaultVal As String) As String
    Dim s As String
    s = Trim$(CStr(wsIn.Range(cellAddr).Value))
    If s = "" Then s = defaultVal
    ReadText = s
End Function

' Regresses the ticker's monthly returns on the benchmark's and stores the
' resulting beta into the params dictionary. Called once per run, not per row.
Public Sub AttachBeta(ByVal params As Object, ByVal ticker As String)
    Dim b As Variant
    b = ComputeBeta(ticker, CStr(params("indexSymbol")), CLng(params("betaMonths")))
    If IsNumeric(b) Then
        params("beta") = CDbl(b)
        params("betaIsEstimated") = True
        Exit Sub
    End If
    params("beta") = 1#
    params("betaIsEstimated") = False
End Sub

' Monthly-return beta = Cov(stock, index) / Var(index).
' Monthly rather than daily returns deliberately: daily returns on a thinly
' traded small cap are dominated by non-synchronous trading, which biases beta
' toward zero. Returns "" when fewer than 12 overlapping months are available.
Public Function ComputeBeta(ByVal ticker As String, ByVal indexSymbol As String, ByVal months As Long) As Variant
    ComputeBeta = ""
    If Trim$(ticker) = "" Then Exit Function
    If Trim$(indexSymbol) = "" Then Exit Function

    Dim fromDate As Date
    fromDate = DateAdd("m", -(months + 1), Date)

    Dim pStock As Object, pIndex As Object
    On Error Resume Next
    Set pStock = FetchDailyPrices(ticker, fromDate, Date)
    Set pIndex = FetchDailyPrices(indexSymbol, fromDate, Date)
    On Error GoTo 0
    If pStock Is Nothing Then Exit Function
    If pIndex Is Nothing Then Exit Function
    If pStock.Count = 0 Then Exit Function
    If pIndex.Count = 0 Then Exit Function

    Dim mStock As Object, mIndex As Object
    Set mStock = MonthEndCloses(pStock)
    Set mIndex = MonthEndCloses(pIndex)

    ' Months present in both series, oldest first.
    Dim keys As Collection
    Set keys = SortedCommonKeys(mStock, mIndex)
    If keys.Count < 13 Then Exit Function

    Dim rs() As Double, rm() As Double
    ReDim rs(1 To keys.Count - 1)
    ReDim rm(1 To keys.Count - 1)

    Dim i As Long, used As Long
    used = 0
    For i = 2 To keys.Count
        Dim s0 As Double, s1 As Double, m0 As Double, m1 As Double
        s0 = CDbl(mStock(CStr(keys(i - 1))))
        s1 = CDbl(mStock(CStr(keys(i))))
        m0 = CDbl(mIndex(CStr(keys(i - 1))))
        m1 = CDbl(mIndex(CStr(keys(i))))
        If s0 > 0 Then
            If m0 > 0 Then
                used = used + 1
                rs(used) = s1 / s0 - 1
                rm(used) = m1 / m0 - 1
            End If
        End If
    Next i
    If used < 12 Then Exit Function

    Dim sumS As Double, sumM As Double
    For i = 1 To used
        sumS = sumS + rs(i)
        sumM = sumM + rm(i)
    Next i
    Dim avgS As Double, avgM As Double
    avgS = sumS / used
    avgM = sumM / used

    Dim cov As Double, varM As Double
    For i = 1 To used
        cov = cov + (rs(i) - avgS) * (rm(i) - avgM)
        varM = varM + (rm(i) - avgM) ^ 2
    Next i
    If varM = 0 Then Exit Function

    ComputeBeta = cov / varM
End Function

' Collapses a daily "yyyy-mm-dd" -> close dictionary into "yyyy-mm" -> last
' close of that month.
Private Function MonthEndCloses(ByVal daily As Object) As Object
    Dim out As Object, lastDay As Object
    Set out = CreateObject("Scripting.Dictionary")
    Set lastDay = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In daily.Keys
        Dim ds As String, ym As String
        ds = CStr(k)
        If Len(ds) >= 7 Then
            ym = Left$(ds, 7)
            Dim isNewer As Boolean
            isNewer = True
            If lastDay.Exists(ym) Then
                If CStr(lastDay(ym)) >= ds Then isNewer = False
            End If
            If isNewer Then
                lastDay(ym) = ds
                out(ym) = daily(k)
            End If
        End If
    Next k

    Set MonthEndCloses = out
End Function

' Keys present in both dictionaries, ascending. Insertion sort -- the series is
' at most a few hundred months, so nothing fancier earns its complexity.
Private Function SortedCommonKeys(ByVal a As Object, ByVal b As Object) As Collection
    Dim out As Collection
    Set out = New Collection

    Dim n As Long
    n = 0
    Dim arr() As String

    Dim k As Variant
    For Each k In a.Keys
        If b.Exists(CStr(k)) Then
            n = n + 1
            ReDim Preserve arr(1 To n)
            arr(n) = CStr(k)
        End If
    Next k
    If n = 0 Then
        Set SortedCommonKeys = out
        Exit Function
    End If

    Dim i As Long, j As Long
    Dim tmp As String
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j) <= tmp Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i

    For i = 1 To n
        out.Add arr(i)
    Next i
    Set SortedCommonKeys = out
End Function

' ---------- Per-row metric block ----------

' Computes columns 49-58 for one filing row and writes them.
'
' Every input is a Variant that may be "" (tag absent from this filing). The
' guards below are deliberately written as nested Ifs rather than
' `If IsNumeric(x) And CDbl(x) > 0`, because VBA's And does NOT short-circuit:
' it would evaluate the CDbl on a non-numeric "" and raise a type mismatch.
'
' periodsPerYear is 1 for an annual row and 4 for a quarterly one; it scales the
' flow figures (NOPAT, sustainable growth) up to an annual rate so they are
' comparable with WACC. Use PeriodsPerYear() on the row's own form label.
Public Sub WriteValuationCells(ByVal ws As Worksheet, ByVal rowIdx As Long, _
    ByVal advCurrentV As Variant, ByVal advNoncurrentV As Variant, _
    ByVal netIncomeV As Variant, ByVal equityV As Variant, ByVal epsV As Variant, ByVal dpsV As Variant, _
    ByVal assetsV As Variant, ByVal currentLiabV As Variant, _
    ByVal operatingIncomeV As Variant, ByVal taxRateV As Variant, _
    ByVal sharesV As Variant, ByVal priceV As Variant, _
    ByVal shortTermDebtV As Variant, ByVal longTermDebtV As Variant, ByVal interestExpenseV As Variant, _
    ByVal params As Object, ByVal periodsPerYear As Double, _
    ByVal dpsTagExists As Boolean, ByVal advSource As String, ByVal dpsIsAnnual As Boolean)

    Dim advTotal As Variant
    advTotal = AddIfAny(advCurrentV, advNoncurrentV)

    ' --- Sustainable growth = ROE x retention ---
    Dim retention As Variant, roe As Variant, sustG As Variant
    ' A Taiwanese dividend is declared once for a whole fiscal year, so on a
    ' quarterly row it must be weighed against a full year of earnings, not one
    ' quarter's -- otherwise a normal payer looks like it is distributing four
    ' times its profit and the retention ratio floors at zero. SEC filers tag a
    ' per-quarter declared dividend, so there the two already line up.
    Dim epsForPayout As Variant
    epsForPayout = epsV
    If dpsIsAnnual Then
        If IsNumeric(epsV) Then epsForPayout = CDbl(epsV) * periodsPerYear
    End If
    retention = RetentionRatio(epsForPayout, dpsV, dpsTagExists)
    roe = Divide(netIncomeV, equityV)
    sustG = ""
    If IsNumeric(roe) Then
        If IsNumeric(retention) Then
            sustG = CDbl(roe) * periodsPerYear * CDbl(retention)
        End If
    End If

    ' --- Invested capital / NOPAT / ROIC ---
    Dim investedCapital As Variant
    investedCapital = Subtract(assetsV, currentLiabV)

    Dim taxRate As Double
    taxRate = CDbl(params("taxFallback"))
    If IsNumeric(taxRateV) Then
        Dim tr As Double
        tr = CDbl(taxRateV)
        ' A negative or absurd effective rate (a tax-benefit quarter, a one-off
        ' true-up) would flip NOPAT's sign or wipe it out; fall back to the
        ' Input-sheet rate rather than propagate it.
        If tr >= 0 Then
            If tr < 0.6 Then taxRate = tr
        End If
    End If

    Dim nopat As Variant
    nopat = ""
    If IsNumeric(operatingIncomeV) Then
        nopat = CDbl(operatingIncomeV) * (1 - taxRate) * periodsPerYear
    End If

    Dim roic As Variant
    roic = ""
    If IsNumeric(nopat) Then
        If IsNumeric(investedCapital) Then
            If CDbl(investedCapital) > 0 Then roic = CDbl(nopat) / CDbl(investedCapital)
        End If
    End If

    ' --- WACC (CAPM) ---
    Dim wacc As Variant
    wacc = ComputeWacc(equityV, sharesV, priceV, shortTermDebtV, longTermDebtV, interestExpenseV, _
        taxRate, params, periodsPerYear)

    Dim spread As Variant
    spread = ""
    If IsNumeric(roic) Then
        If IsNumeric(wacc) Then spread = CDbl(roic) - CDbl(wacc)
    End If

    With ws
        .Cells(rowIdx, COL_ADV_CURRENT).Value = advCurrentV
        .Cells(rowIdx, COL_ADV_NONCURRENT).Value = advNoncurrentV
        .Cells(rowIdx, COL_ADV_TOTAL).Value = advTotal
        .Cells(rowIdx, COL_RETENTION).Value = retention
        .Cells(rowIdx, COL_SUSTAINABLE_G).Value = sustG
        .Cells(rowIdx, COL_INVESTED_CAPITAL).Value = investedCapital
        .Cells(rowIdx, COL_NOPAT).Value = nopat
        .Cells(rowIdx, COL_ROIC).Value = roic
        .Cells(rowIdx, COL_WACC).Value = wacc
        .Cells(rowIdx, COL_SPREAD).Value = spread
        .Cells(rowIdx, COL_ADV_SOURCE).Value = advSource
    End With
End Sub

' WACC = E/(D+E) x Ke + D/(D+E) x Kd x (1 - t), with Ke from CAPM.
'
' E prefers MARKET value (shares x price at the report date); it degrades to
' book equity when the share count or the price is missing, which is common for
' multi-class filers (see modSEC's mapShares comment) and for TW tickers whose
' market suffix didn't resolve. That degradation changes the weighting, not the
' validity, so it is silent rather than blanking the whole column.
Public Function ComputeWacc(ByVal equityV As Variant, ByVal sharesV As Variant, ByVal priceV As Variant, _
    ByVal shortTermDebtV As Variant, ByVal longTermDebtV As Variant, ByVal interestExpenseV As Variant, _
    ByVal taxRate As Double, ByVal params As Object, ByVal periodsPerYear As Double) As Variant

    ComputeWacc = ""

    Dim e As Double
    e = -1
    If IsNumeric(sharesV) Then
        If IsNumeric(priceV) Then
            If CDbl(sharesV) > 0 Then
                If CDbl(priceV) > 0 Then e = CDbl(sharesV) * CDbl(priceV)
            End If
        End If
    End If
    If e <= 0 Then
        If IsNumeric(equityV) Then
            If CDbl(equityV) > 0 Then e = CDbl(equityV)
        End If
    End If
    If e <= 0 Then Exit Function

    Dim d As Double
    d = 0
    If IsNumeric(shortTermDebtV) Then
        If CDbl(shortTermDebtV) > 0 Then d = d + CDbl(shortTermDebtV)
    End If
    If IsNumeric(longTermDebtV) Then
        If CDbl(longTermDebtV) > 0 Then d = d + CDbl(longTermDebtV)
    End If

    Dim rf As Double, erp As Double, beta As Double
    rf = CDbl(params("rf"))
    erp = CDbl(params("erp"))
    beta = CDbl(params("beta"))

    Dim ke As Double
    ke = rf + beta * erp

    If d <= 0 Then
        ComputeWacc = ke
        Exit Function
    End If

    ' Cost of debt from the filing's own interest expense, annualised the same
    ' way NOPAT is. A period that repaid its debt mid-quarter still carries
    ' interest expense against a near-zero closing balance and can produce an
    ' absurd implied rate, so the result is clamped to a plausible band before
    ' use; outside it, fall back to Rf + 200bp.
    Dim kd As Double
    kd = rf + 0.02
    If IsNumeric(interestExpenseV) Then
        If CDbl(interestExpenseV) > 0 Then
            Dim implied As Double
            implied = CDbl(interestExpenseV) * periodsPerYear / d
            If implied > 0.001 Then
                If implied < 0.35 Then kd = implied
            End If
        End If
    End If

    Dim total As Double
    total = d + e
    ComputeWacc = (e / total) * ke + (d / total) * kd * (1 - taxRate)
End Function

'  1 - payout ratio, from per-share figures. Returns "" when EPS is missing or
' non-positive: a loss period makes the payout ratio meaningless, and a company
' paying a dividend out of a loss would otherwise show a retention ratio above
' 1, implying growth funded by nothing.
'
' dpsTagExists distinguishes "this filer paid no dividend" from "this data
' source has no dividend tag at all". SEC filings tag dividends per share, so
' an absent value there genuinely means a non-payer and full retention (pass
' True). MOPS's quarterly statements carry no reliable per-share dividend tag
' whatsoever (see modMOPS.ExtractTWMetrics), so treating absence as zero there
' would silently report retention = 100% -- i.e. sustainable growth = ROE -- for
' every Taiwanese company including the reliable dividend payers. Pass False and
' the column stays blank instead of confidently wrong.
Public Function RetentionRatio(ByVal epsV As Variant, ByVal dpsV As Variant, ByVal dpsTagExists As Boolean) As Variant
    RetentionRatio = ""
    If Not IsNumeric(epsV) Then Exit Function
    If CDbl(epsV) <= 0 Then Exit Function

    Dim dps As Double
    dps = 0
    If IsNumeric(dpsV) Then
        If CDbl(dpsV) > 0 Then dps = CDbl(dpsV)
    ElseIf Not dpsTagExists Then
        Exit Function
    End If

    Dim r As Double
    r = 1 - dps / CDbl(epsV)
    If r < 0 Then r = 0      ' dividend exceeded EPS this period -- no retention
    RetentionRatio = r
End Function

' 1 for an annual filing, 4 for a quarterly one.
Public Function PeriodsPerYear(ByVal form As String, ByVal annualFormMarker As String) As Double
    If InStr(1, form, annualFormMarker, vbTextCompare) > 0 Then
        PeriodsPerYear = 1
    Else
        PeriodsPerYear = 4
    End If
End Function

' ---------- small numeric helpers (mirrors of modSEC's SafeXxx) ----------

Private Function AddIfAny(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim aOk As Boolean, bOk As Boolean
    aOk = IsNumeric(a)
    bOk = IsNumeric(b)
    If aOk Then
        If bOk Then
            AddIfAny = CDbl(a) + CDbl(b)
        Else
            AddIfAny = CDbl(a)
        End If
    ElseIf bOk Then
        AddIfAny = CDbl(b)
    Else
        AddIfAny = ""
    End If
End Function

Private Function Divide(ByVal numer As Variant, ByVal denom As Variant) As Variant
    Divide = ""
    If Not IsNumeric(numer) Then Exit Function
    If Not IsNumeric(denom) Then Exit Function
    If CDbl(denom) = 0 Then Exit Function
    Divide = CDbl(numer) / CDbl(denom)
End Function

Private Function Subtract(ByVal a As Variant, ByVal b As Variant) As Variant
    Subtract = ""
    If Not IsNumeric(a) Then Exit Function
    If Not IsNumeric(b) Then Exit Function
    Subtract = CDbl(a) - CDbl(b)
End Function

' Applies the number formats for columns 49-58 across a written block.
Public Sub FormatValuationColumns(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    If lastRow < firstRow Then Exit Sub
    ws.Range(ws.Cells(firstRow, COL_ADV_CURRENT), ws.Cells(lastRow, COL_ADV_TOTAL)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(firstRow, COL_RETENTION), ws.Cells(lastRow, COL_SUSTAINABLE_G)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(firstRow, COL_INVESTED_CAPITAL), ws.Cells(lastRow, COL_NOPAT)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(firstRow, COL_ROIC), ws.Cells(lastRow, COL_SPREAD)).NumberFormat = "0.00%"
End Sub
