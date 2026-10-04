Attribute VB_Name = "modSECData"
Option Explicit

' ============================================================================
'  modSECData  --  headless SEC EDGAR / XBRL financial-data reader
' ----------------------------------------------------------------------------
'  Shared module (repo: shared-vba/). Imported by BOTH the standalone
'  SECFilingFetcher workbook AND the RR4 "Company research" integration.
'
'  Purpose: expose the fetch-and-parse core of modSEC.FetchSECFilings WITHOUT
'  any dependency on the Input / Filings sheets, on modCharts, or on modTheme,
'  so a caller that just wants the numbers (RR4's Company research lower half)
'  can get them as plain Dictionary/Collection data and render them itself.
'
'  Depends only on:  modHttp.HttpGet , modJsonUtil.*   (both in shared-vba/)
'
'  Public surface (deliberately tiny, to avoid any name clash when this module
'  sits in the same VBAProject as the full modSEC.bas):
'    GetXbrlFinancials(tickerOrCik, wantAnnual, wantQuarter) As Object
'    XbrlValue(maps, conceptKey, accn, reportDate, targetForm) As Variant
'    XbrlFyFp(allMaps, accn, reportDate, targetForm) As String
'
'  Everything else is Private. The Private helpers (ResolveCIK, GetFilings,
'  BuildConceptMap, MergeConceptMaps*, LookupConceptValue, LookupFyFp, ...)
'  are verbatim copies of the same-named routines in modSEC.bas as of the P0
'  audit -- kept Private here so importing this module alongside modSEC in the
'  standalone workbook raises no "Ambiguous name detected". modSEC.bas is NOT
'  modified by this file; converging the two (trimming modSEC to call into
'  this module) is a later, separately-verified step.
' ============================================================================


' ---------------------------------------------------------------------------
'  PUBLIC: one call -> everything a renderer needs
' ---------------------------------------------------------------------------
' Returns a Scripting.Dictionary:
'   "ok"          Boolean   -- False on any failure (never raises to caller)
'   "error"       String    -- populated when ok = False
'   "cik"         String    -- zero-padded 10-digit CIK
'   "ticker"      String
'   "entityName"  String
'   "filings10K"  Collection of filing dicts  (form/amended/accn/filingDate/
'   "filings10Q"  Collection of filing dicts   reportDate/primaryDocument/items)
'   "maps"        Dictionary  conceptKey -> (accn -> Collection of raw fact JSON)
'                 keys: Revenue GrossProfit OperatingIncome NetIncome Eps
'                       Assets Liabilities Cfo Rd Sga Dividends Inventory Ar
'                       CurrentAssets CurrentLiabilities LongTermDebt
'                       ShortTermDebt StockholdersEquity EffectiveTaxRate
'                       CapEx Cash Da Cogs InterestExpense Shares AccountsPayable
'   "allMaps"     Variant() -- (Assets,NetIncome,Revenue,Liabilities,Cfo,Eps)
'                 fallback chain for XbrlFyFp
'
' Renderer usage: walk res("filings10Q") newest-first; per filing f and per
'   concept k:  v = XbrlValue(res("maps"), k, f("accn"), f("reportDate"), f("form"))
'   period label:  XbrlFyFp(res("allMaps"), f("accn"), f("reportDate"), f("form"))
Public Function GetXbrlFinancials(ByVal tickerOrCik As String, ByVal wantAnnual As Long, ByVal wantQuarter As Long) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res("ok") = False
    res("error") = ""

    On Error GoTo EH

    Dim inputText As String
    inputText = Trim$(tickerOrCik)
    If inputText = "" Then
        res("error") = "empty ticker/CIK"
        Set GetXbrlFinancials = res
        Exit Function
    End If

    Dim wantK As Long, wantQ As Long
    wantK = wantAnnual
    If wantK < 1 Then wantK = 7
    wantQ = wantQuarter
    If wantQ < 1 Then wantQ = 10

    Dim cik As String, tickerTitle As String, ticker As String
    If Not ResolveCIK(inputText, cik, tickerTitle, ticker) Then
        res("error") = "cannot resolve CIK for: " & inputText
        Set GetXbrlFinancials = res
        Exit Function
    End If

    Dim filings10K As Collection, filings10Q As Collection
    Dim filings8K As Collection, filingsForm4 As Collection
    Call GetFilings(cik, wantK, wantQ, 0, 0, filings10K, filings10Q, filings8K, filingsForm4)

    If filings10K.Count = 0 And filings10Q.Count = 0 Then
        res("error") = "company found (CIK " & cik & ") but no 10-K/10-Q filings"
        Set GetXbrlFinancials = res
        Exit Function
    End If

    Dim companyFactsJson As String
    companyFactsJson = HttpGet("https://data.sec.gov/api/xbrl/companyfacts/CIK" & cik & ".json")

    Dim entityName As String
    entityName = ExtractJsonString(companyFactsJson, "entityName")
    ' SEC pads some names with no-break spaces (U+00A0); make them plain
    ' spaces so the name compares, searches and exports like normal text.
    entityName = Replace(entityName, ChrW(160), " ")
    If entityName = "" Then entityName = tickerTitle

    Dim factsRaw As String, usgaapRaw As String, deiRaw As String
    factsRaw = ExtractJsonValueRaw(companyFactsJson, "facts")
    usgaapRaw = ExtractJsonValueRaw(factsRaw, "us-gaap")
    deiRaw = ExtractJsonValueRaw(factsRaw, "dei")

    Dim maps As Object
    Set maps = CreateObject("Scripting.Dictionary")

    Set maps("Revenue") = BuildConceptMap(usgaapRaw, Array("Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"), "USD")
    Set maps("NetIncome") = BuildConceptMap(usgaapRaw, Array("NetIncomeLoss", "ProfitLoss"), "USD")
    Set maps("Eps") = BuildConceptMap(usgaapRaw, Array("EarningsPerShareDiluted", "EarningsPerShareBasicAndDiluted"), "USD/shares")
    Set maps("Assets") = BuildConceptMap(usgaapRaw, Array("Assets"), "USD")
    Set maps("Liabilities") = BuildConceptMap(usgaapRaw, Array("Liabilities"), "USD")
    Set maps("Cfo") = BuildConceptMap(usgaapRaw, Array("NetCashProvidedByUsedInOperatingActivities", "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"), "USD")

    ' Point-in-time shares outstanding: multi-class filers (Class A/B common
    ' stock) tag it PER CLASS with a ClassOfStockAxis dimension, never as one
    ' undimensioned company-wide fact, and the companyfacts API only exposes
    ' undimensioned facts -- so both instant tags come back empty for them.
    ' Last-resort fallback: the EPS-denominator weighted-average share count
    ' (basic preferred; a PERIOD AVERAGE, not a period-END count).
    ' MergeConceptMapsFallback (gap-fill only) is required, not MergeConceptMaps
    ' (blend): the real instant fact has no "start", so blending would let the
    ' weighted-average fallback outrank a correct real count on every filer
    ' where the primary tags DO work.
    Set maps("Shares") = MergeConceptMapsFallback( _
        MergeConceptMaps(BuildConceptMap(deiRaw, Array("EntityCommonStockSharesOutstanding"), "shares"), BuildConceptMap(usgaapRaw, Array("CommonStockSharesOutstanding"), "shares")), _
        BuildConceptMap(usgaapRaw, Array("WeightedAverageNumberOfSharesOutstandingBasic", "WeightedAverageNumberOfDilutedSharesOutstanding"), "shares"))

    Set maps("GrossProfit") = BuildConceptMap(usgaapRaw, Array("GrossProfit"), "USD")
    Set maps("OperatingIncome") = BuildConceptMap(usgaapRaw, Array("OperatingIncomeLoss"), "USD")
    Set maps("Rd") = BuildConceptMap(usgaapRaw, Array("ResearchAndDevelopmentExpense"), "USD")
    Set maps("Sga") = BuildConceptMap(usgaapRaw, Array("SellingGeneralAndAdministrativeExpense"), "USD")
    Set maps("Dividends") = BuildConceptMap(usgaapRaw, Array("CommonStockDividendsPerShareDeclared", "CommonStockDividendsPerShareCashPaid"), "USD/shares")
    Set maps("Inventory") = BuildConceptMap(usgaapRaw, Array("InventoryNet"), "USD")
    Set maps("Ar") = BuildConceptMap(usgaapRaw, Array("AccountsReceivableNetCurrent"), "USD")
    Set maps("CurrentAssets") = BuildConceptMap(usgaapRaw, Array("AssetsCurrent"), "USD")
    Set maps("CurrentLiabilities") = BuildConceptMap(usgaapRaw, Array("LiabilitiesCurrent"), "USD")

    ' Fallback for a filer that stops tagging the current/noncurrent debt split
    ' and reports one combined DebtInstrumentCarryingAmount total instead.
    ' Gap-fill only -- never overrides the primary split tags when they exist.
    ' 2026-10-04: MU stopped tagging LongTermDebt after FY2026Q1 and reports only
    ' LongTermDebtAndCapitalLeaseObligations (debt + finance leases, a wider
    ' basis), so that is a last-resort gap-fill after the two above.
    Set maps("LongTermDebt") = MergeConceptMapsFallback( _
        MergeConceptMapsFallback( _
            BuildConceptMap(usgaapRaw, Array("LongTermDebtNoncurrent", "LongTermDebt"), "USD"), _
            BuildConceptMap(usgaapRaw, Array("DebtInstrumentCarryingAmount"), "USD")), _
        BuildConceptMap(usgaapRaw, Array("LongTermDebtAndCapitalLeaseObligations"), "USD"))

    Set maps("StockholdersEquity") = BuildConceptMap(usgaapRaw, Array("StockholdersEquity"), "USD")
    Set maps("EffectiveTaxRate") = BuildConceptMap(usgaapRaw, Array("EffectiveIncomeTaxRateContinuingOperations"), "pure")
    Set maps("CapEx") = BuildConceptMap(usgaapRaw, Array("PaymentsToAcquirePropertyPlantAndEquipment"), "USD")
    Set maps("Cash") = BuildConceptMap(usgaapRaw, Array("CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"), "USD")
    Set maps("Da") = BuildConceptMap(usgaapRaw, Array("DepreciationDepletionAndAmortization", "DepreciationAmortizationAndAccretionNet", "DepreciationAndAmortization"), "USD")
    Set maps("Cogs") = BuildConceptMap(usgaapRaw, Array("CostOfGoodsAndServicesSold", "CostOfRevenue", "CostOfGoodsSold"), "USD")
    Set maps("InterestExpense") = BuildConceptMap(usgaapRaw, Array("InterestExpense", "InterestExpenseDebt"), "USD")
    Set maps("ShortTermDebt") = BuildConceptMap(usgaapRaw, Array("ShortTermBorrowings", "DebtCurrent", "LongTermDebtCurrent"), "USD")
    Set maps("AccountsPayable") = BuildConceptMap(usgaapRaw, Array("AccountsPayableCurrent", "AccountsPayableTradeCurrent"), "USD")

    Dim allMaps As Variant
    allMaps = Array(maps("Assets"), maps("NetIncome"), maps("Revenue"), maps("Liabilities"), maps("Cfo"), maps("Eps"))

    res("ok") = True
    res("error") = ""
    res("cik") = cik
    res("ticker") = ticker
    res("entityName") = entityName
    Set res("filings10K") = filings10K
    Set res("filings10Q") = filings10Q
    Set res("maps") = maps
    res("allMaps") = allMaps

    Set GetXbrlFinancials = res
    Exit Function

EH:
    res("ok") = False
    res("error") = "GetXbrlFinancials error: " & Err.Description & " [src: " & Err.Source & "]"
    Set GetXbrlFinancials = res
End Function


' ---------------------------------------------------------------------------
'  PUBLIC: thin readers over the private lookups (renamed to avoid clashing
'  with modSEC.LookupConceptValue / modSEC.LookupFyFp in the same project)
' ---------------------------------------------------------------------------
Public Function XbrlValue(ByVal maps As Object, ByVal conceptKey As String, ByVal accn As String, ByVal reportDate As String, ByVal targetForm As String) As Variant
    If maps Is Nothing Then
        XbrlValue = ""
        Exit Function
    End If
    If Not maps.Exists(conceptKey) Then
        XbrlValue = ""
        Exit Function
    End If
    XbrlValue = LookupConceptValue(maps(conceptKey), accn, reportDate, targetForm)
End Function

Public Function XbrlFyFp(ByVal allMaps As Variant, ByVal accn As String, ByVal reportDate As String, ByVal targetForm As String) As String
    XbrlFyFp = LookupFyFp(allMaps, accn, reportDate, targetForm)
End Function


' ===========================================================================
'  PRIVATE  --  verbatim from modSEC.bas (P0 audit snapshot). Do not edit here
'  in isolation; keep in sync with modSEC.bas until the convergence step.
' ===========================================================================

Private Function RegexEscape(ByVal s As String) As String
    Dim specials As String
    specials = "\^$.|?*+()[]{}"
    Dim result As String
    Dim i As Long
    For i = 1 To Len(s)
        Dim c As String
        c = Mid$(s, i, 1)
        If InStr(specials, c) > 0 Then
            result = result & "\" & c
        Else
            result = result & c
        End If
    Next i
    RegexEscape = result
End Function

Private Function ResolveCIK(ByVal inputText As String, ByRef cik As String, ByRef title As String, ByRef ticker As String) As Boolean
    Dim raw As String
    raw = HttpGet("https://www.sec.gov/files/company_tickers.json")

    Dim q As String
    q = Chr(34)

    Dim upperInput As String
    upperInput = RegexEscape(UCase$(Trim$(inputText)))

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = q & "cik_str" & q & ":(\d+)," & q & "ticker" & q & ":" & q & upperInput & q & "," & q & "title" & q & ":" & q & "([^" & q & "]*)" & q

    Dim mc As Object
    Set mc = re.Execute(raw)
    If mc.Count > 0 Then
        cik = Format$(CLng(mc(0).SubMatches(0)), "0000000000")
        title = mc(0).SubMatches(1)
        ticker = UCase$(Trim$(inputText))
        ResolveCIK = True
        Exit Function
    End If

    Dim nameEscaped As String
    nameEscaped = RegexEscape(Trim$(inputText))

    re.IgnoreCase = True
    re.Pattern = q & "cik_str" & q & ":(\d+)," & q & "ticker" & q & ":" & q & "([^" & q & "]+)" & q & "," & q & "title" & q & ":" & q & "([^" & q & "]*" & nameEscaped & "[^" & q & "]*)" & q

    Set mc = re.Execute(raw)
    If mc.Count > 0 Then
        cik = Format$(CLng(mc(0).SubMatches(0)), "0000000000")
        ticker = mc(0).SubMatches(1)
        title = mc(0).SubMatches(2)
        ResolveCIK = True
        Exit Function
    End If

    cik = ""
    title = ""
    ticker = ""
    ResolveCIK = False
End Function

Private Sub GetFilings(ByVal cik As String, ByVal wantK As Long, ByVal wantQ As Long, ByVal want8K As Long, ByVal want4 As Long, _
    ByRef filings10K As Collection, ByRef filings10Q As Collection, ByRef filings8K As Collection, ByRef filingsForm4 As Collection)

    Set filings10K = New Collection
    Set filings10Q = New Collection
    Set filings8K = New Collection
    Set filingsForm4 = New Collection

    ' Keyed by reportDate so that when a 10-K/A (or 10-Q/A) amends an earlier
    ' original for the same fiscal period, only the amendment is kept. EDGAR's
    ' submissions.json is scanned newest-filed-first, and an amendment is
    ' always filed after what it amends -- so "first reportDate seen wins"
    ' naturally keeps the amendment when one exists.
    Dim seenReportDatesK As Object, seenReportDatesQ As Object
    Set seenReportDatesK = CreateObject("Scripting.Dictionary")
    Set seenReportDatesQ = CreateObject("Scripting.Dictionary")

    Dim json As String
    json = HttpGet("https://data.sec.gov/submissions/CIK" & cik & ".json")

    Dim filingsRaw As String
    filingsRaw = ExtractJsonValueRaw(json, "filings")

    Dim recentRaw As String
    recentRaw = ExtractJsonValueRaw(filingsRaw, "recent")

    Call ScanRecentBlock(recentRaw, wantK, wantQ, want8K, want4, filings10K, filings10Q, filings8K, filingsForm4, seenReportDatesK, seenReportDatesQ)

    If filings10K.Count < wantK Or filings10Q.Count < wantQ Or filings8K.Count < want8K Or filingsForm4.Count < want4 Then
        Dim filesRaw As String
        filesRaw = ExtractJsonValueRaw(filingsRaw, "files")
        Dim fileElems As Collection
        Set fileElems = SplitJsonArrayElements(filesRaw)

        Dim fe As Variant
        For Each fe In fileElems
            If filings10K.Count >= wantK And filings10Q.Count >= wantQ And filings8K.Count >= want8K And filingsForm4.Count >= want4 Then Exit For
            Dim fname As String
            fname = ExtractJsonString(CStr(fe), "name")
            If fname <> "" Then
                Dim oldJson As String
                oldJson = HttpGet("https://data.sec.gov/submissions/" & fname)
                Call ScanRecentBlock(oldJson, wantK, wantQ, want8K, want4, filings10K, filings10Q, filings8K, filingsForm4, seenReportDatesK, seenReportDatesQ)
            End If
        Next fe
    End If
End Sub

Private Sub ScanRecentBlock(ByRef block As String, ByVal wantK As Long, ByVal wantQ As Long, ByVal want8K As Long, ByVal want4 As Long, _
    ByRef filings10K As Collection, ByRef filings10Q As Collection, ByRef filings8K As Collection, ByRef filingsForm4 As Collection, _
    ByRef seenReportDatesK As Object, ByRef seenReportDatesQ As Object)

    Dim formsRaw As String, accnRaw As String, filingDateRaw As String, reportDateRaw As String, primDocRaw As String, itemsRaw As String
    formsRaw = ExtractJsonValueRaw(block, "form")
    accnRaw = ExtractJsonValueRaw(block, "accessionNumber")
    filingDateRaw = ExtractJsonValueRaw(block, "filingDate")
    reportDateRaw = ExtractJsonValueRaw(block, "reportDate")
    primDocRaw = ExtractJsonValueRaw(block, "primaryDocument")
    itemsRaw = ExtractJsonValueRaw(block, "items")

    Dim forms As Collection, accns As Collection, fdates As Collection, rdates As Collection, pdocs As Collection, itemsColl As Collection
    Set forms = SplitJsonArrayElements(formsRaw)
    Set accns = SplitJsonArrayElements(accnRaw)
    Set fdates = SplitJsonArrayElements(filingDateRaw)
    Set rdates = SplitJsonArrayElements(reportDateRaw)
    Set pdocs = SplitJsonArrayElements(primDocRaw)
    Set itemsColl = SplitJsonArrayElements(itemsRaw)

    Dim n As Long
    n = forms.Count
    Dim i As Long
    For i = 1 To n
        If filings10K.Count >= wantK And filings10Q.Count >= wantQ And filings8K.Count >= want8K And filingsForm4.Count >= want4 Then Exit For
        Dim f As String
        f = JsonUnescape(CStr(forms(i)))
        Dim rd As String
        rd = ""
        If i <= rdates.Count Then rd = JsonUnescape(CStr(rdates(i)))
        If (f = "10-K" Or f = "10-K/A") And filings10K.Count < wantK And (rd = "" Or Not seenReportDatesK.Exists(rd)) Then
            filings10K.Add BuildFilingDict(f, accns, fdates, rdates, pdocs, itemsColl, i, (f = "10-K/A"))
            If rd <> "" Then seenReportDatesK(rd) = True
        ElseIf (f = "10-Q" Or f = "10-Q/A") And filings10Q.Count < wantQ And (rd = "" Or Not seenReportDatesQ.Exists(rd)) Then
            filings10Q.Add BuildFilingDict(f, accns, fdates, rdates, pdocs, itemsColl, i, (f = "10-Q/A"))
            If rd <> "" Then seenReportDatesQ(rd) = True
        ElseIf f = "8-K" And filings8K.Count < want8K Then
            filings8K.Add BuildFilingDict(f, accns, fdates, rdates, pdocs, itemsColl, i, False)
        ElseIf f = "4" And filingsForm4.Count < want4 Then
            filingsForm4.Add BuildFilingDict(f, accns, fdates, rdates, pdocs, itemsColl, i, False)
        End If
    Next i
End Sub

Private Function BuildFilingDict(ByVal f As String, ByRef accns As Collection, ByRef fdates As Collection, ByRef rdates As Collection, ByRef pdocs As Collection, ByRef itemsColl As Collection, ByVal idx As Long, ByVal isAmended As Boolean) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("form") = f
    d("amended") = isAmended
    If idx <= accns.Count Then d("accn") = JsonUnescape(CStr(accns(idx))) Else d("accn") = ""
    If idx <= fdates.Count Then d("filingDate") = JsonUnescape(CStr(fdates(idx))) Else d("filingDate") = ""
    If idx <= rdates.Count Then d("reportDate") = JsonUnescape(CStr(rdates(idx))) Else d("reportDate") = ""
    If idx <= pdocs.Count Then d("primaryDocument") = JsonUnescape(CStr(pdocs(idx))) Else d("primaryDocument") = ""
    If idx <= itemsColl.Count Then d("items") = JsonUnescape(CStr(itemsColl(idx))) Else d("items") = ""
    Set BuildFilingDict = d
End Function

Private Function BuildConceptMap(ByRef usgaapRaw As String, ByVal tagCandidates As Variant, ByVal unitKey As String) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim t As Variant
    For Each t In tagCandidates
        Dim conceptRaw As String
        conceptRaw = ExtractJsonValueRaw(usgaapRaw, CStr(t))
        If Len(conceptRaw) > 0 Then
            Dim unitsRaw As String
            unitsRaw = ExtractJsonValueRaw(conceptRaw, "units")
            Dim arrRaw As String
            arrRaw = ExtractJsonValueRaw(unitsRaw, unitKey)
            If Len(arrRaw) > 0 Then
                Dim elems As Collection
                Set elems = SplitJsonArrayElements(arrRaw)
                Dim e As Variant
                For Each e In elems
                    Dim accn As String
                    accn = ExtractJsonString(CStr(e), "accn")
                    If accn <> "" Then
                        If Not map.Exists(accn) Then
                            ' "Dim col As New Collection" here is the classic VBA trap:
                            ' As-New only auto-instantiates when the variable is
                            ' currently Nothing, and Dim inside a loop is hoisted to
                            ' one procedure-scoped variable -- so after the FIRST accn
                            ' creates a Collection, every SUBSEQUENT distinct accn would
                            ' silently reuse that same live reference instead of getting
                            ' its own, merging every accn's facts into one shared
                            ' Collection. Set col = New Collection forces a genuinely
                            ' fresh instance every time this branch runs.
                            Dim col As Collection
                            Set col = New Collection
                            map.Add accn, col
                        End If
                        map(accn).Add CStr(e)
                    End If
                Next e
            End If
        End If
    Next t

    Set BuildConceptMap = map
End Function

' Combines two accn->Collection maps built from different XBRL taxonomy sections
' (e.g. dei vs us-gaap) so a lookup can fall back from one source to the other.
Private Function MergeConceptMaps(ByVal a As Object, ByVal b As Object) As Object
    Dim merged As Object
    Set merged = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    If Not a Is Nothing Then
        For Each k In a.Keys
            Set merged(k) = a(k)
        Next k
    End If
    If Not b Is Nothing Then
        For Each k In b.Keys
            If Not merged.Exists(k) Then
                Set merged(k) = b(k)
            Else
                Dim e As Variant
                For Each e In b(k)
                    merged(k).Add e
                Next e
            End If
        Next k
    End If

    Set MergeConceptMaps = merged
End Function

' Like MergeConceptMaps, but for an accn already present in `primary`, the
' fallback map's facts for that accn are discarded rather than blended in.
' Only an accn the primary map has NO entry for at all receives the fallback's
' facts.
Private Function MergeConceptMapsFallback(ByVal primary As Object, ByVal fallback As Object) As Object
    Dim merged As Object
    Set merged = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    If Not primary Is Nothing Then
        For Each k In primary.Keys
            Set merged(k) = primary(k)
        Next k
    End If
    If Not fallback Is Nothing Then
        For Each k In fallback.Keys
            If Not merged.Exists(k) Then Set merged(k) = fallback(k)
        Next k
    End If

    Set MergeConceptMapsFallback = merged
End Function

' For flow concepts (revenue, net income, cash flow), 10-Qs commonly tag BOTH the
' standalone quarter AND the year-to-date cumulative figure under the same "end"
' date, distinguished only by "start". We want the quarter-only value, which is
' the one with the LATEST (shortest-period) start date. Balance-sheet concepts
' (Assets, Liabilities) have no "start" at all, so this is a no-op for them.
Private Function LookupConceptValue(ByVal map As Object, ByVal accn As String, ByVal reportDate As String, ByVal targetForm As String) As Variant
    If map Is Nothing Then
        LookupConceptValue = ""
        Exit Function
    End If
    If Not map.Exists(accn) Then
        LookupConceptValue = ""
        Exit Function
    End If

    Dim col As Collection
    Set col = map(accn)

    Dim best As String, bestStart As String
    Dim e As Variant
    For Each e In col
        If ExtractJsonString(CStr(e), "end") = reportDate And ExtractJsonString(CStr(e), "form") = targetForm Then
            Dim s As String
            s = ExtractJsonString(CStr(e), "start")
            If best = "" Or s > bestStart Then
                best = CStr(e)
                bestStart = s
            End If
        End If
    Next e

    If best = "" Then
        For Each e In col
            If ExtractJsonString(CStr(e), "end") = reportDate Then
                Dim s2 As String
                s2 = ExtractJsonString(CStr(e), "start")
                If best = "" Or s2 > bestStart Then
                    best = CStr(e)
                    bestStart = s2
                End If
            End If
        Next e
    End If

    If best = "" Then best = CStr(col(col.Count))

    Dim valRaw As String
    valRaw = ExtractJsonValueRaw(best, "val")
    LookupConceptValue = valRaw
End Function

' A filing's XBRL facts include prior-period comparatives under the same accession
' number (e.g. a 10-K shows 3 years of income statement history), all sharing the
' same accn -- so we must match on "end" (and ideally "form") to avoid picking up
' a comparative figure from a different fiscal period.
Private Function LookupFyFp(ByVal maps As Variant, ByVal accn As String, ByVal reportDate As String, ByVal targetForm As String) As String
    Dim m As Variant
    For Each m In maps
        If Not m Is Nothing Then
            If m.Exists(accn) Then
                Dim col As Collection
                Set col = m(accn)

                Dim best As String
                Dim e As Variant
                For Each e In col
                    If ExtractJsonString(CStr(e), "end") = reportDate And ExtractJsonString(CStr(e), "form") = targetForm Then
                        best = CStr(e)
                        Exit For
                    End If
                Next e
                If best = "" Then
                    For Each e In col
                        If ExtractJsonString(CStr(e), "end") = reportDate Then
                            best = CStr(e)
                            Exit For
                        End If
                    Next e
                End If
                If best = "" Then best = CStr(col(col.Count))

                Dim fy As String, fp As String
                fy = ExtractJsonString(best, "fy")
                fp = ExtractJsonString(best, "fp")
                If fy <> "" Then
                    LookupFyFp = fy & "|" & fp
                    Exit Function
                End If
            End If
        End If
    Next m
    LookupFyFp = "|"
End Function
