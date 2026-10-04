Attribute VB_Name = "modMOPSData"
Option Explicit

' ============================================================================
'  modMOPSData  --  headless Taiwan MOPS (mopsov.twse.com.tw) financial reader
' ----------------------------------------------------------------------------
'  Shared module (repo: shared-vba/). TW counterpart of modSECData.
'
'  Purpose: expose the fetch-and-parse core of modMOPS WITHOUT any dependency
'  on the Input sheet, on modCharts, or on modTheme, so a caller that just
'  wants the numbers (RR4's Company research lower half) can get them as plain
'  Dictionary/Collection data and render them itself.
'
'  Depends only on:  modPrices.FetchDailyPrices  (via FetchTwPrices, optional)
'                    -- plus WinHttp / ADODB.Stream directly (HttpGetBig5).
'  No dependency on modHttp: MOPS pages are Big5 with no charset header and
'  need the ADODB.Stream override; this module carries its own HttpGetBig5.
'
'  Public surface (tiny, Private-everything-else, so it can sit in the same
'  VBAProject as the full modMOPS with no "Ambiguous name detected"):
'    GetTwFinancials(coId, quarters) As Object
'    GetTwMonthlyRevenue(coId, months) As Object
'
'  Private helpers are verbatim from modMOPS.bas as of the P0 audit (CJK-only
'  comments rewritten to English so this file stays pure ASCII -- see the
'  vba-bas-ascii-only memory). modMOPS.bas is NOT modified by this file;
'  converging the two is a later, separately-verified step.
' ============================================================================

Private Const MOPS_USER_AGENT As String = "Ryan Personal Research Tool ryan929929@gmail.com"


' ---------------------------------------------------------------------------
'  PUBLIC: quarterly + annual financials, one call
' ---------------------------------------------------------------------------
' Walks quarters backward from ~95 days ago until it has `quarters` quarterly
' periods (Q4 also yields an annual/FY period). Returns a Scripting.Dictionary:
'   "ok"                Boolean   -- False on any failure (never raises)
'   "error"             String
'   "coId"              String
'   "entityName"        String    -- Chinese company name from the iXBRL, or coId
'   "filingsQuarterly"  Collection of filing dicts (accn/form/reportDate/...)
'   "filingsAnnual"     Collection of filing dicts
'   "quarterlyRows"     Collection of row specs -- .Item("metrics") is a dict of
'   "annualRows"        Collection of row specs    the 25 metric keys already
'                                                  extracted for that period
'   "maps"              Dictionary  metricKey -> (accn -> Collection of JSON)
'                       -- same shape modSEC produces; kept for parity. The
'                       renderer normally just walks quarterlyRows/annualRows.
'   "allMaps"           Variant()  (Assets, Revenue, NetIncome)
'
' Metric keys (in each row's "metrics" dict and in "maps"):
'   Revenue GrossProfit OperatingIncome NetIncome Eps Assets Liabilities Cfo
'   Rd Sga Inventory Ar CurrentAssets CurrentLiabilities LongTermDebt
'   ShortTermDebt StockholdersEquity CapEx Cash Da Cogs InterestExpense
'   EffectiveTaxRate Shares AccountsPayable
' Amounts are in whole NTD: t164sb01 iXBRL tags carry a scale attribute
' (usually 3) which ParseIxFacts already applies. (The separate monthly-revenue
' feed in GetTwMonthlyRevenue IS in thousands -- see its note.)
Public Function GetTwFinancials(ByVal coId As String, ByVal quarters As Long) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res("ok") = False
    res("error") = ""

    On Error GoTo EH

    Dim cid As String
    cid = Trim$(coId)
    If cid = "" Then
        res("error") = "empty CO_ID"
        Set GetTwFinancials = res
        Exit Function
    End If

    Dim wantQ As Long
    wantQ = quarters
    If wantQ < 1 Then wantQ = 10

    ' Back off 95 days so the first quarter tried is essentially always already
    ' published (Q1-Q3 reviewed reports land ~45 days after quarter-end, the
    ' Q4/annual audited report ~75-90 days).
    Dim curYear As Long, curSeason As Long
    Call QuarterOf(DateAdd("d", -95, Date), curYear, curSeason)

    Dim filingsAnnual As New Collection, filingsQuarterly As New Collection
    Dim annualRows As New Collection, quarterlyRows As New Collection
    Dim maps As Object
    Set maps = NewMapSet()

    Dim entityName As String
    entityName = ""

    Dim gotQ As Long
    gotQ = 0
    Dim triesLeft As Long
    triesLeft = wantQ + 8

    Do While gotQ < wantQ And triesLeft > 0
        triesLeft = triesLeft - 1

        Dim usedUrl As String, fetchedName As String
        Dim facts As Object
        Set facts = FetchQuarterFacts(cid, curYear, curSeason, usedUrl, fetchedName)

        If facts.Count > 0 Then
            If entityName = "" Then entityName = fetchedName

            Dim qStart As Date, qEnd As Date, fyStart As Date
            Call QuarterBounds(curYear, curSeason, qStart, qEnd, fyStart)

            Dim accnQ As String, fyLabel As String, fpLabelQ As String
            fyLabel = CStr(curYear)
            fpLabelQ = "Q" & curSeason
            accnQ = "TW" & cid & "-" & fyLabel & fpLabelQ

            Dim metricsQ As Object
            Set metricsQ = ExtractTWMetrics(facts, qStart, qEnd, qEnd)
            Call AccumulateMetricsIntoMaps(maps, metricsQ, accnQ, "MOPS-Q", fyLabel, fpLabelQ, qStart, qEnd, qEnd)
            filingsQuarterly.Add BuildTWFilingDict(accnQ, "MOPS-Q", qEnd)
            quarterlyRows.Add BuildTWRowSpec(entityName, cid, metricsQ, "MOPS-Q", fyLabel, fpLabelQ, qEnd, usedUrl)
            gotQ = gotQ + 1

            If curSeason = 4 Then
                Dim accnA As String, fpLabelA As String
                fpLabelA = "FY"
                accnA = "TW" & cid & "-" & fyLabel & fpLabelA

                Dim metricsA As Object
                Set metricsA = ExtractTWMetrics(facts, fyStart, qEnd, qEnd)
                Call AccumulateMetricsIntoMaps(maps, metricsA, accnA, "MOPS-10-K", fyLabel, fpLabelA, fyStart, qEnd, qEnd)
                filingsAnnual.Add BuildTWFilingDict(accnA, "MOPS-10-K", qEnd)
                annualRows.Add BuildTWRowSpec(entityName, cid, metricsA, "MOPS-10-K", fyLabel, fpLabelA, qEnd, usedUrl)
            End If
        End If

        curSeason = curSeason - 1
        If curSeason = 0 Then
            curSeason = 4
            curYear = curYear - 1
        End If
    Loop

    If filingsQuarterly.Count = 0 Then
        res("error") = "no MOPS iXBRL financial data found for CO_ID " & cid
        Set GetTwFinancials = res
        Exit Function
    End If
    If entityName = "" Then entityName = cid

    Dim allMapsArr As Variant
    allMapsArr = Array(maps("Assets"), maps("Revenue"), maps("NetIncome"))

    res("ok") = True
    res("error") = ""
    res("coId") = cid
    res("entityName") = entityName
    Set res("filingsQuarterly") = filingsQuarterly
    Set res("filingsAnnual") = filingsAnnual
    Set res("quarterlyRows") = quarterlyRows
    Set res("annualRows") = annualRows
    Set res("maps") = maps
    res("allMaps") = allMapsArr

    Set GetTwFinancials = res
    Exit Function

EH:
    res("ok") = False
    res("error") = "GetTwFinancials error: " & Err.Description & " [src: " & Err.Source & "]"
    Set GetTwFinancials = res
End Function


' ---------------------------------------------------------------------------
'  PUBLIC: monthly revenue trend (MOPS t21sc03), fetch + recompute only
' ---------------------------------------------------------------------------
' Returns a Scripting.Dictionary:
'   "ok"      Boolean
'   "error"   String
'   "coId"    String
'   "labels"  Variant() 1..n of "yyyy-mm" (oldest first)
'   "rev"     Variant() 1..n single-month revenue (thousands NTD), recomputed
'             as (this month cumulative - prior month cumulative)
'   "yoy"     Variant() 1..n year-over-year ratio (or Empty when no comparator)
Public Function GetTwMonthlyRevenue(ByVal coId As String, ByVal months As Long) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res("ok") = False
    res("error") = ""

    On Error GoTo EH

    Dim cid As String
    cid = Trim$(coId)
    If cid = "" Then
        res("error") = "empty CO_ID"
        Set GetTwMonthlyRevenue = res
        Exit Function
    End If

    Dim want As Long
    want = months
    If want < 1 Then want = 24

    Dim results As Object
    Set results = CreateObject("Scripting.Dictionary")   ' "yyyy-mm" -> Array(thisRevRaw, cumRev)

    Dim cursor As Date
    cursor = DateAdd("m", -1, Date)   ' this calendar month's revenue isn't published until next month

    Dim fetched As Long
    fetched = 0
    Dim triesLeft As Long
    triesLeft = want + 3

    Do While fetched < want And triesLeft > 0
        triesLeft = triesLeft - 1

        Dim rocYear As Long, mo As Long
        rocYear = Year(cursor) - 1911
        mo = Month(cursor)

        Dim thisRev As Variant, cumRev As Variant
        Dim okRow As Boolean
        okRow = FetchOneMonthRevenue(cid, rocYear, mo, "sii", thisRev, cumRev)
        If Not okRow Then okRow = FetchOneMonthRevenue(cid, rocYear, mo, "otc", thisRev, cumRev)

        If okRow Then
            results(Format$(cursor, "yyyy-mm")) = Array(thisRev, cumRev)
            fetched = fetched + 1
        End If

        cursor = DateAdd("m", -1, cursor)
    Loop

    If results.Count = 0 Then
        res("error") = "no monthly-revenue rows found for CO_ID " & cid
        Set GetTwMonthlyRevenue = res
        Exit Function
    End If

    Dim n As Long
    n = results.Count
    Dim keys() As String
    ReDim keys(1 To n)
    Dim idx As Long
    idx = 0
    Dim k As Variant
    For Each k In results.Keys
        idx = idx + 1
        keys(idx) = CStr(k)
    Next k

    ' Insertion sort ascending. NOT "Do While b >= 1 And keys(b) > tmpK" --
    ' VBA's And does not short-circuit, so that form indexes keys(0) out of
    ' bounds once b reaches 0.
    Dim a As Long, b As Long, tmpK As String
    For a = 2 To n
        tmpK = keys(a)
        b = a - 1
        Do While b >= 1
            If keys(b) <= tmpK Then Exit Do
            keys(b + 1) = keys(b)
            b = b - 1
        Loop
        keys(b + 1) = tmpK
    Next a

    Dim labelArr() As Variant, revArr() As Variant, yoyArr() As Variant
    ReDim labelArr(1 To n): ReDim revArr(1 To n): ReDim yoyArr(1 To n)

    For idx = 1 To n
        Dim curKey As String
        curKey = keys(idx)
        labelArr(idx) = curKey

        Dim curMonthNum As Long
        curMonthNum = CLng(Right$(curKey, 2))
        Dim curCum As Double
        curCum = results(curKey)(1)

        If curMonthNum = 1 Then
            revArr(idx) = results(curKey)(0)
        Else
            Dim priorKey As String
            priorKey = Format$(DateAdd("m", -1, YMToDate(curKey)), "yyyy-mm")
            If results.Exists(priorKey) Then
                revArr(idx) = curCum - results(priorKey)(1)
            Else
                revArr(idx) = results(curKey)(0)
            End If
        End If
    Next idx

    For idx = 1 To n
        Dim priorYearKey As String
        priorYearKey = Format$(DateAdd("yyyy", -1, YMToDate(keys(idx))), "yyyy-mm")
        If results.Exists(priorYearKey) Then
            Dim priorYearRev As Variant
            Dim pyMonthNum As Long
            pyMonthNum = CLng(Right$(priorYearKey, 2))
            If pyMonthNum = 1 Then
                priorYearRev = results(priorYearKey)(0)
            Else
                Dim pyPriorKey As String
                pyPriorKey = Format$(DateAdd("m", -1, YMToDate(priorYearKey)), "yyyy-mm")
                If results.Exists(pyPriorKey) Then
                    priorYearRev = results(priorYearKey)(1) - results(pyPriorKey)(1)
                Else
                    priorYearRev = results(priorYearKey)(0)
                End If
            End If
            If IsNumeric(priorYearRev) And CDbl(priorYearRev) <> 0 Then
                yoyArr(idx) = revArr(idx) / CDbl(priorYearRev) - 1
            Else
                yoyArr(idx) = Empty
            End If
        Else
            yoyArr(idx) = Empty
        End If
    Next idx

    res("ok") = True
    res("error") = ""
    res("coId") = cid
    res("labels") = labelArr
    res("rev") = revArr
    res("yoy") = yoyArr

    Set GetTwMonthlyRevenue = res
    Exit Function

EH:
    res("ok") = False
    res("error") = "GetTwMonthlyRevenue error: " & Err.Description & " [src: " & Err.Source & "]"
    Set GetTwMonthlyRevenue = res
End Function


' ===========================================================================
'  PRIVATE  --  verbatim from modMOPS.bas (P0 audit snapshot; CJK comments
'  rewritten to English). Keep in sync with modMOPS.bas until convergence.
' ===========================================================================

Private Function MetricKeys() As Variant
    MetricKeys = Array("Revenue", "GrossProfit", "OperatingIncome", "NetIncome", "Eps", "Assets", "Liabilities", "Cfo", _
        "Rd", "Sga", "Inventory", "Ar", "CurrentAssets", "CurrentLiabilities", "LongTermDebt", "ShortTermDebt", _
        "StockholdersEquity", "CapEx", "Cash", "Da", "Cogs", "InterestExpense", "EffectiveTaxRate", "Shares", "AccountsPayable")
End Function

' Fetches one (CO_ID, SYEAR, SSEASON) quarterly report, trying REPORT_ID=C
' (consolidated) first, falling back to A (individual) when C has zero facts.
Private Function FetchQuarterFacts(ByVal coId As String, ByVal sYear As Long, ByVal sSeason As Long, ByRef usedUrlOut As String, ByRef entityNameOut As String) As Object
    Dim url As String
    url = MopsReportUrl(coId, sYear, sSeason, "C")
    Dim html As String
    html = HttpGetBig5(url)
    Dim facts As Object
    Set facts = ParseIxFacts(html)

    If facts.Count = 0 Then
        url = MopsReportUrl(coId, sYear, sSeason, "A")
        html = HttpGetBig5(url)
        Set facts = ParseIxFacts(html)
    End If

    If facts.Count > 0 Then
        entityNameOut = ExtractIxNonNumeric(html, "tifrs-notes:CompanyChineseName")
    End If

    usedUrlOut = url
    Set FetchQuarterFacts = facts
End Function

Private Function MopsReportUrl(ByVal coId As String, ByVal sYear As Long, ByVal sSeason As Long, ByVal reportId As String) As String
    MopsReportUrl = "https://mopsov.twse.com.tw/server-java/t164sb01?step=1&CO_ID=" & coId & "&SYEAR=" & sYear & "&SSEASON=" & sSeason & "&REPORT_ID=" & reportId
End Function

Private Sub QuarterOf(ByVal d As Date, ByRef y As Long, ByRef q As Long)
    y = Year(d)
    q = ((Month(d) - 1) \ 3) + 1
End Sub

Private Sub QuarterBounds(ByVal y As Long, ByVal q As Long, ByRef qStart As Date, ByRef qEnd As Date, ByRef fyStart As Date)
    Dim startMonth As Long
    startMonth = (q - 1) * 3 + 1
    qStart = DateSerial(y, startMonth, 1)
    qEnd = DateSerial(y, startMonth + 3, 0)
    fyStart = DateSerial(y, 1, 1)
End Sub

' Scans raw MOPS iXBRL HTML for every <ix:nonFraction name contextRef ...>value
' </ix:nonFraction> tag -> Dictionary keyed "conceptName|contextRef" -> numeric
' value (comma-stripped, scale-multiplied, sign-applied). Skips dimensional
' contexts (contextRef containing "_").
Private Function ParseIxFacts(ByRef html As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If Len(html) = 0 Then
        Set ParseIxFacts = d
        Exit Function
    End If

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<ix:nonfraction\s+([^>]*)>([^<]*)</ix:nonfraction>"

    Dim mc As Object
    Set mc = re.Execute(html)

    Dim m As Object
    For Each m In mc
        Dim attrs As String, raw As String
        attrs = m.SubMatches(0)
        raw = Trim$(m.SubMatches(1))

        Dim conceptName As String, contextRef As String
        conceptName = ExtractAttr(attrs, "name")
        contextRef = ExtractAttr(attrs, "contextRef")
        If conceptName <> "" And contextRef <> "" And InStr(contextRef, "_") = 0 Then
            Dim numStr As String
            numStr = Replace(raw, ",", "")
            If IsNumeric(numStr) Then
                Dim v As Double
                v = CDbl(numStr)

                Dim scaleStr As String
                scaleStr = ExtractAttr(attrs, "scale")
                If IsNumeric(scaleStr) Then v = v * (10 ^ CLng(scaleStr))

                If ExtractAttr(attrs, "sign") = "-" Then v = -v

                d(conceptName & "|" & contextRef) = v
            End If
        End If
    Next m

    Set ParseIxFacts = d
End Function

' Pulls attrName="value" out of an ix: tag's attribute string. Case-insensitive
' and order-independent.
Private Function ExtractAttr(ByRef attrs As String, ByVal attrName As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = attrName & "\s*=\s*""([^""]*)"""
    Dim mc As Object
    Set mc = re.Execute(attrs)
    If mc.Count > 0 Then
        ExtractAttr = mc(0).SubMatches(0)
    Else
        ExtractAttr = ""
    End If
End Function

Private Function ExtractIxNonNumeric(ByRef html As String, ByVal conceptName As String) As String
    If Len(html) = 0 Then Exit Function
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "<ix:nonnumeric\s+name=""" & conceptName & """[^>]*>([^<]*)</ix:nonnumeric>"
    Dim mc As Object
    Set mc = re.Execute(html)
    If mc.Count > 0 Then
        ExtractIxNonNumeric = Trim$(mc(0).SubMatches(0))
    Else
        ExtractIxNonNumeric = ""
    End If
End Function

' Locale-independent "yyyy-mm" -> Date (first of month). CDate would rely on
' Windows regional settings for date-part order.
Private Function YMToDate(ByVal ym As String) As Date
    YMToDate = DateSerial(CLng(Left$(ym, 4)), CLng(Mid$(ym, 6, 2)), 1)
End Function

Private Function CtxInstant(ByVal d As Date) As String
    CtxInstant = "AsOf" & Format$(d, "yyyymmdd")
End Function

Private Function CtxDuration(ByVal startD As Date, ByVal endD As Date) As String
    CtxDuration = "From" & Format$(startD, "yyyymmdd") & "To" & Format$(endD, "yyyymmdd")
End Function

' Looks up conceptName|contextRef for the first matching candidate name.
Private Function GetFact(ByVal facts As Object, ByVal candidates As Variant, ByVal contextRef As String) As Variant
    Dim c As Variant
    For Each c In candidates
        Dim key As String
        key = CStr(c) & "|" & contextRef
        If facts.Exists(key) Then
            GetFact = facts(key)
            Exit Function
        End If
    Next c
    GetFact = ""
End Function

' Reads every tracked metric for ONE period out of an already-parsed facts dict
' (durStart..durEnd for flow concepts, instEnd for point-in-time concepts).
Private Function ExtractTWMetrics(ByVal facts As Object, ByVal durStart As Date, ByVal durEnd As Date, ByVal instEnd As Date) As Object
    Dim instCtx As String, durCtx As String
    instCtx = CtxInstant(instEnd)
    durCtx = CtxDuration(durStart, durEnd)

    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")

    m("Revenue") = GetFact(facts, Array("ifrs-full:Revenue"), durCtx)
    m("GrossProfit") = GetFact(facts, Array("ifrs-full:GrossProfit"), durCtx)
    m("OperatingIncome") = GetFact(facts, Array("ifrs-full:ProfitLossFromOperatingActivities"), durCtx)
    m("NetIncome") = GetFact(facts, Array("ifrs-full:ProfitLossAttributableToOwnersOfParent"), durCtx)

    Dim epsV As Variant
    epsV = GetFact(facts, Array("ifrs-full:DilutedEarningsLossPerShare"), durCtx)
    If Not (IsNumeric(epsV) And CStr(epsV) <> "") Then epsV = GetFact(facts, Array("ifrs-full:BasicEarningsLossPerShare"), durCtx)
    m("Eps") = epsV

    m("Assets") = GetFact(facts, Array("ifrs-full:Assets"), instCtx)
    m("Liabilities") = GetFact(facts, Array("ifrs-full:Liabilities"), instCtx)
    m("Cfo") = GetFact(facts, Array("ifrs-full:CashFlowsFromUsedInOperatingActivities", "ifrs-full:CashFlowsFromUsedInOperations"), durCtx)
    m("Rd") = GetFact(facts, Array("ifrs-full:ResearchAndDevelopmentExpense"), durCtx)
    m("Sga") = SumIfNumeric(GetFact(facts, Array("ifrs-full:SellingExpense"), durCtx), GetFact(facts, Array("ifrs-full:AdministrativeExpense"), durCtx))
    m("Inventory") = GetFact(facts, Array("ifrs-full:Inventories"), instCtx)
    m("Ar") = GetFact(facts, Array("tifrs-bsci-ci:AccountsReceivableNet"), instCtx)
    m("CurrentAssets") = GetFact(facts, Array("ifrs-full:CurrentAssets"), instCtx)
    m("CurrentLiabilities") = GetFact(facts, Array("ifrs-full:CurrentLiabilities"), instCtx)
    m("LongTermDebt") = GetFact(facts, Array("ifrs-full:LongtermBorrowings"), instCtx)
    m("ShortTermDebt") = GetFact(facts, Array("ifrs-full:ShorttermBorrowings"), instCtx)
    m("StockholdersEquity") = GetFact(facts, Array("ifrs-full:EquityAttributableToOwnersOfParent"), instCtx)

    ' TW cash-flow items are SIGNED (capex carries sign="-"); the shared FCF
    ' formula expects the positive "amount paid" convention, so normalize to a
    ' magnitude here.
    Dim capExRaw As Variant
    capExRaw = GetFact(facts, Array("ifrs-full:PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities"), durCtx)
    If IsNumeric(capExRaw) And CStr(capExRaw) <> "" Then capExRaw = Abs(CDbl(capExRaw))
    m("CapEx") = capExRaw
    m("Cash") = GetFact(facts, Array("ifrs-full:CashAndCashEquivalents"), instCtx)

    ' TW reconciles depreciation and amortisation as two separate line items;
    ' summed here so the rest of the pipeline sees one D&A number.
    m("Da") = SumIfNumeric(GetFact(facts, Array("ifrs-full:AdjustmentsForDepreciationExpense"), durCtx), GetFact(facts, Array("ifrs-full:AdjustmentsForAmortisationExpense"), durCtx))
    m("Cogs") = GetFact(facts, Array("tifrs-bsci-ci:OperatingCosts"), durCtx)
    m("AccountsPayable") = GetFact(facts, Array("ifrs-full:TradeAndOtherCurrentPayablesToTradeSuppliers"), instCtx)
    m("InterestExpense") = GetFact(facts, Array("ifrs-full:FinanceCosts"), durCtx)

    ' No direct effective-tax-rate tag in the TW taxonomy -- computed from the
    ' two income-statement lines that are tagged.
    m("EffectiveTaxRate") = SafeDivVariant(GetFact(facts, Array("ifrs-full:IncomeTaxExpenseContinuingOperations"), durCtx), GetFact(facts, Array("ifrs-full:ProfitLossBeforeTax", "tifrs-SCF:ProfitLossBeforeTax"), durCtx))

    ' No shares-outstanding tag on this endpoint -- backed out from
    ' diluted EPS = NetIncome / dilutedWeightedShares.
    m("Shares") = SafeDivVariant(m("NetIncome"), epsV)

    Set ExtractTWMetrics = m
End Function

Private Function SumIfNumeric(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim aOk As Boolean, bOk As Boolean
    aOk = IsNumeric(a) And CStr(a) <> ""
    bOk = IsNumeric(b) And CStr(b) <> ""
    If aOk And bOk Then
        SumIfNumeric = CDbl(a) + CDbl(b)
    ElseIf aOk Then
        SumIfNumeric = CDbl(a)
    ElseIf bOk Then
        SumIfNumeric = CDbl(b)
    Else
        SumIfNumeric = ""
    End If
End Function

Private Function SafeDivVariant(ByVal numer As Variant, ByVal denom As Variant) As Variant
    If IsNumeric(numer) And CStr(numer) <> "" And IsNumeric(denom) And CStr(denom) <> "" Then
        If CDbl(denom) <> 0 Then
            SafeDivVariant = CDbl(numer) / CDbl(denom)
            Exit Function
        End If
    End If
    SafeDivVariant = ""
End Function

Private Function NewMapSet() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In MetricKeys()
        d.Add CStr(k), CreateObject("Scripting.Dictionary")
    Next k
    Set NewMapSet = d
End Function

' Feeds one period's metrics into mapSet's per-metric accn->Collection maps, as
' small hand-built JSON objects carrying the fields modSEC.LookupConceptValue /
' LookupFyFp read. Str(v) not CStr/Format$ for the number literal: those render
' with the system locale's decimal separator; Str() always uses ".".
Private Sub AccumulateMetricsIntoMaps(ByVal mapSet As Object, ByVal metrics As Object, ByVal accn As String, ByVal form As String, ByVal fy As String, ByVal fp As String, ByVal startD As Date, ByVal endD As Date, ByVal instEnd As Date)
    Dim k As Variant
    For Each k In MetricKeys()
        Dim v As Variant
        v = metrics(CStr(k))
        If IsNumeric(v) And CStr(v) <> "" Then
            Dim useStart As Date
            If CStr(k) = "Assets" Or CStr(k) = "Liabilities" Or CStr(k) = "Inventory" Or CStr(k) = "Ar" Or _
               CStr(k) = "CurrentAssets" Or CStr(k) = "CurrentLiabilities" Or CStr(k) = "LongTermDebt" Or _
               CStr(k) = "ShortTermDebt" Or CStr(k) = "StockholdersEquity" Or CStr(k) = "Cash" Or CStr(k) = "Shares" Or _
               CStr(k) = "AccountsPayable" Then
                useStart = instEnd   ' point-in-time concepts: start==end
            Else
                useStart = startD
            End If

            Dim entry As String
            entry = "{""accn"":""" & accn & """,""form"":""" & form & """,""fy"":""" & fy & """,""fp"":""" & fp & _
                """,""start"":""" & Format$(useStart, "yyyy-mm-dd") & """,""end"":""" & Format$(endD, "yyyy-mm-dd") & _
                """,""val"":" & Trim$(Str(CDbl(v))) & "}"

            Dim map As Object
            Set map = mapSet(CStr(k))
            If Not map.Exists(accn) Then
                Dim col As Collection
                Set col = New Collection
                map.Add accn, col
            End If
            map(accn).Add entry
        End If
    Next k
End Sub

Private Function BuildTWFilingDict(ByVal accn As String, ByVal form As String, ByVal reportDate As Date) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("accn") = accn
    d("form") = form
    d("reportDate") = Format$(reportDate, "yyyy-mm-dd")
    d("filingDate") = Format$(reportDate, "yyyy-mm-dd")
    d("primaryDocument") = ""
    d("items") = ""
    d("amended") = False
    Set BuildTWFilingDict = d
End Function

Private Function BuildTWRowSpec(ByVal entityName As String, ByVal coId As String, ByVal metrics As Object, ByVal form As String, ByVal fyLabel As String, ByVal fpLabel As String, ByVal reportDate As Date, ByVal sourceUrl As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("entityName") = entityName
    d("coId") = coId
    Set d("metrics") = metrics
    d("form") = form
    d("fyLabel") = fyLabel
    d("fpLabel") = fpLabel
    d("reportDate") = reportDate
    d("sourceUrl") = sourceUrl
    Set BuildTWRowSpec = d
End Function

' Parses one month's MOPS company-level revenue table
' (t21sc03_<ROCyear>_<month>.html, market "sii" listed / "otc" OTC) for coId.
' Only reads the two positionally-safe columns: this-month revenue (index 0)
' and cumulative revenue (index 5).
Private Function FetchOneMonthRevenue(ByVal coId As String, ByVal rocYear As Long, ByVal mo As Long, ByVal market As String, ByRef thisRev As Variant, ByRef cumRev As Variant) As Boolean
    FetchOneMonthRevenue = False

    Dim url As String
    url = "https://mopsov.twse.com.tw/nas/t21/" & market & "/t21sc03_" & rocYear & "_" & mo & ".html"

    Dim html As String
    html = HttpGetBig5(url)
    If Len(html) = 0 Then Exit Function

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "<td[^>]*>\s*" & coId & "\s*</td>\s*<td[^>]*>[^<]*</td>((?:\s*<td[^>]*>[^<]*</td>){6,9})"

    Dim mc As Object
    Set mc = re.Execute(html)
    If mc.Count = 0 Then Exit Function

    Dim reCell As Object
    Set reCell = CreateObject("VBScript.RegExp")
    reCell.Global = True
    reCell.IgnoreCase = True
    reCell.Pattern = "<td[^>]*>([^<]*)</td>"

    Dim cellMc As Object
    Set cellMc = reCell.Execute(CStr(mc(0).SubMatches(0)))
    If cellMc.Count < 6 Then Exit Function

    Dim thisRaw As String, cumRaw As String
    thisRaw = Replace(Trim$(cellMc(0).SubMatches(0)), ",", "")
    cumRaw = Replace(Trim$(cellMc(5).SubMatches(0)), ",", "")

    If Not IsNumeric(thisRaw) Or Not IsNumeric(cumRaw) Then Exit Function

    thisRev = CDbl(thisRaw)
    cumRev = CDbl(cumRaw)
    FetchOneMonthRevenue = True
End Function

' MOPS serves Big5 HTML with no charset in the Content-Type header. WinHttp's
' .ResponseText would decode with a Western codepage and mangle every
' double-byte char -- go through raw .ResponseBody + an ADODB.Stream with an
' explicit Big5 override instead.
Private Function HttpGetBig5(ByVal url As String) As String
    HttpGetBig5 = ""

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 15000, 60000
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", MOPS_USER_AGENT
    http.Send

    If http.Status <> 200 Then Exit Function

    Dim bytes() As Byte
    bytes = http.ResponseBody

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1   ' adTypeBinary
    stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.Type = 2   ' adTypeText
    stm.Charset = "big5"
    HttpGetBig5 = stm.ReadText
    stm.Close
End Function
