Attribute VB_Name = "modSEC"
Option Explicit

' Entry point. Triggered by the Input sheet's Worksheet_Change when A1 changes.
Public Sub FetchSECFilings()
    Dim wsIn As Worksheet, wsOut As Worksheet
    Set wsIn = ThisWorkbook.Sheets("Input")
    Set wsOut = ThisWorkbook.Sheets("Filings")

    On Error GoTo ErrHandler

    Dim inputText As String
    inputText = Trim$(CStr(wsIn.Range("A1").Value))
    If inputText = "" Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    SetStatus wsIn, "查詢 CIK 中..."

    Dim cik As String, tickerTitle As String, ticker As String
    If Not ResolveCIK(inputText, cik, tickerTitle, ticker) Then
        SetStatus wsIn, "找不到公司，請確認股票代號，或改用更完整的公司名稱關鍵字：" & inputText
        GoTo CleanExit
    End If

    Dim wantK As Long, wantQ As Long, want8K As Long, want4 As Long
    wantK = GetSetting(wsIn, "B4", 7)
    wantQ = GetSetting(wsIn, "B5", 8)
    want8K = GetSetting(wsIn, "B6", 10)
    want4 = GetSetting(wsIn, "B7", 10)

    SetStatus wsIn, "抓取申報清單中 (CIK " & cik & ")..."
    Dim filings10K As Collection, filings10Q As Collection, filings8K As Collection, filingsForm4 As Collection
    Call GetFilings(cik, wantK, wantQ, want8K, want4, filings10K, filings10Q, filings8K, filingsForm4)

    If filings10K.Count = 0 And filings10Q.Count = 0 Then
        SetStatus wsIn, "找到公司 (CIK " & cik & ") 但沒有 10-K/10-Q 資料"
        GoTo CleanExit
    End If

    SetStatus wsIn, "抓取財務數字 (XBRL) 中，資料量較大請稍候..."
    Dim companyFactsJson As String
    companyFactsJson = HttpGet("https://data.sec.gov/api/xbrl/companyfacts/CIK" & cik & ".json")

    Dim entityName As String
    entityName = ExtractJsonString(companyFactsJson, "entityName")
    If entityName = "" Then entityName = tickerTitle

    Dim factsRaw As String, usgaapRaw As String, deiRaw As String
    factsRaw = ExtractJsonValueRaw(companyFactsJson, "facts")
    usgaapRaw = ExtractJsonValueRaw(factsRaw, "us-gaap")
    deiRaw = ExtractJsonValueRaw(factsRaw, "dei")

    Dim mapRevenue As Object, mapNetIncome As Object, mapEps As Object
    Dim mapAssets As Object, mapLiabilities As Object, mapCFO As Object, mapShares As Object
    Dim mapGrossProfit As Object, mapOperatingIncome As Object, mapRD As Object, mapSGA As Object, mapDividends As Object
    Dim mapInventory As Object, mapAR As Object, mapCurrentAssets As Object, mapCurrentLiabilities As Object
    Dim mapLongTermDebt As Object, mapStockholdersEquity As Object, mapEffectiveTaxRate As Object, mapCapEx As Object
    Dim mapCash As Object, mapDA As Object
    Dim mapCOGS As Object, mapInterestExpense As Object
    Dim mapShortTermDebt As Object

    Set mapRevenue = BuildConceptMap(usgaapRaw, Array("Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"), "USD")
    Set mapNetIncome = BuildConceptMap(usgaapRaw, Array("NetIncomeLoss", "ProfitLoss"), "USD")
    Set mapEps = BuildConceptMap(usgaapRaw, Array("EarningsPerShareDiluted", "EarningsPerShareBasicAndDiluted"), "USD/shares")
    Set mapAssets = BuildConceptMap(usgaapRaw, Array("Assets"), "USD")
    Set mapLiabilities = BuildConceptMap(usgaapRaw, Array("Liabilities"), "USD")
    Set mapCFO = BuildConceptMap(usgaapRaw, Array("NetCashProvidedByUsedInOperatingActivities", "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"), "USD")
    Set mapShares = MergeConceptMaps(BuildConceptMap(deiRaw, Array("EntityCommonStockSharesOutstanding"), "shares"), BuildConceptMap(usgaapRaw, Array("CommonStockSharesOutstanding"), "shares"))
    Set mapGrossProfit = BuildConceptMap(usgaapRaw, Array("GrossProfit"), "USD")
    Set mapOperatingIncome = BuildConceptMap(usgaapRaw, Array("OperatingIncomeLoss"), "USD")
    Set mapRD = BuildConceptMap(usgaapRaw, Array("ResearchAndDevelopmentExpense"), "USD")
    Set mapSGA = BuildConceptMap(usgaapRaw, Array("SellingGeneralAndAdministrativeExpense"), "USD")
    Set mapDividends = BuildConceptMap(usgaapRaw, Array("CommonStockDividendsPerShareDeclared", "CommonStockDividendsPerShareCashPaid"), "USD/shares")
    Set mapInventory = BuildConceptMap(usgaapRaw, Array("InventoryNet"), "USD")
    Set mapAR = BuildConceptMap(usgaapRaw, Array("AccountsReceivableNetCurrent"), "USD")
    Set mapCurrentAssets = BuildConceptMap(usgaapRaw, Array("AssetsCurrent"), "USD")
    Set mapCurrentLiabilities = BuildConceptMap(usgaapRaw, Array("LiabilitiesCurrent"), "USD")
    Set mapLongTermDebt = BuildConceptMap(usgaapRaw, Array("LongTermDebtNoncurrent", "LongTermDebt"), "USD")
    Set mapStockholdersEquity = BuildConceptMap(usgaapRaw, Array("StockholdersEquity"), "USD")
    Set mapEffectiveTaxRate = BuildConceptMap(usgaapRaw, Array("EffectiveIncomeTaxRateContinuingOperations"), "pure")
    Set mapCapEx = BuildConceptMap(usgaapRaw, Array("PaymentsToAcquirePropertyPlantAndEquipment"), "USD")
    Set mapCash = BuildConceptMap(usgaapRaw, Array("CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"), "USD")
    Set mapDA = BuildConceptMap(usgaapRaw, Array("DepreciationDepletionAndAmortization", "DepreciationAmortizationAndAccretionNet", "DepreciationAndAmortization"), "USD")
    Set mapCOGS = BuildConceptMap(usgaapRaw, Array("CostOfGoodsAndServicesSold", "CostOfRevenue", "CostOfGoodsSold"), "USD")
    Set mapInterestExpense = BuildConceptMap(usgaapRaw, Array("InterestExpense", "InterestExpenseDebt"), "USD")
    ' Current portion of debt -- previously missing entirely, which made
    ' Enterprise Value understate true debt (only LongTermDebt(Noncurrent) was
    ' counted). Absence of this tag is treated as "no short-term debt" (0) by
    ' BuildSnapshotTableInto's EV formula, not as "unknown".
    Set mapShortTermDebt = BuildConceptMap(usgaapRaw, Array("ShortTermBorrowings", "DebtCurrent", "LongTermDebtCurrent"), "USD")

    SetStatus wsIn, "寫入工作表中..."
    ' Snapshot the accession-number -> segment-note-URL mapping already on the
    ' sheet BEFORE it gets wiped, so filings we've already seen this run don't
    ' each cost a fresh FindSegmentReportUrl HTTP round-trip below. A filing's
    ' FilingSummary.xml is static once filed, so caching an empty-string result
    ' ("no segment note found") is just as safe as caching a real URL.
    Dim knownSegmentUrls As Object
    Set knownSegmentUrls = ReadExistingSegmentUrls(wsOut)
    Call ClearOutputSheet(wsOut)
    Call WriteHeaders(wsOut)

    Dim allMaps As Variant
    allMaps = Array(mapAssets, mapNetIncome, mapRevenue, mapLiabilities, mapCFO, mapEps)

    Dim rowIdx As Long
    rowIdx = 2

    SetStatus wsIn, "寫入 10-K/10-Q 並查詢部門附註連結中（新申報才需額外查詢）..."
    Dim f As Variant
    For Each f In filings10K
        rowIdx = WriteFilingRow(wsOut, rowIdx, entityName, cik, f, mapRevenue, mapNetIncome, mapEps, mapAssets, mapLiabilities, mapCFO, mapGrossProfit, mapOperatingIncome, mapRD, mapSGA, mapDividends, _
            mapInventory, mapAR, mapCurrentAssets, mapCurrentLiabilities, mapLongTermDebt, mapStockholdersEquity, mapEffectiveTaxRate, mapCapEx, mapShares, allMaps, knownSegmentUrls, _
            mapCOGS, mapInterestExpense, mapCash)
    Next f
    For Each f In filings10Q
        rowIdx = WriteFilingRow(wsOut, rowIdx, entityName, cik, f, mapRevenue, mapNetIncome, mapEps, mapAssets, mapLiabilities, mapCFO, mapGrossProfit, mapOperatingIncome, mapRD, mapSGA, mapDividends, _
            mapInventory, mapAR, mapCurrentAssets, mapCurrentLiabilities, mapLongTermDebt, mapStockholdersEquity, mapEffectiveTaxRate, mapCapEx, mapShares, allMaps, knownSegmentUrls, _
            mapCOGS, mapInterestExpense, mapCash)
    Next f

    If rowIdx > 2 Then
        wsOut.Range(wsOut.Cells(2, 6), wsOut.Cells(rowIdx - 1, 7)).NumberFormat = "yyyy-mm-dd"
        wsOut.Range(wsOut.Cells(2, 10), wsOut.Cells(rowIdx - 1, 11)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 12), wsOut.Cells(rowIdx - 1, 12)).NumberFormat = "0.00"
        wsOut.Range(wsOut.Cells(2, 13), wsOut.Cells(rowIdx - 1, 15)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 17), wsOut.Cells(rowIdx - 1, 18)).NumberFormat = "0.00%"
        wsOut.Range(wsOut.Cells(2, 19), wsOut.Cells(rowIdx - 1, 20)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 21), wsOut.Cells(rowIdx - 1, 21)).NumberFormat = "0.00%"
        wsOut.Range(wsOut.Cells(2, 22), wsOut.Cells(rowIdx - 1, 22)).NumberFormat = "0.00"
        wsOut.Range(wsOut.Cells(2, 25), wsOut.Cells(rowIdx - 1, 28)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 29), wsOut.Cells(rowIdx - 1, 29)).NumberFormat = "0.00"
        wsOut.Range(wsOut.Cells(2, 30), wsOut.Cells(rowIdx - 1, 31)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 32), wsOut.Cells(rowIdx - 1, 32)).NumberFormat = "0.00"
        wsOut.Range(wsOut.Cells(2, 33), wsOut.Cells(rowIdx - 1, 33)).NumberFormat = "0.00%"
        wsOut.Range(wsOut.Cells(2, 34), wsOut.Cells(rowIdx - 1, 35)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 36), wsOut.Cells(rowIdx - 1, 38)).NumberFormat = "0.00%"
        wsOut.Range(wsOut.Cells(2, 39), wsOut.Cells(rowIdx - 1, 45)).NumberFormat = "0.00"
        wsOut.Range(wsOut.Cells(2, 46), wsOut.Cells(rowIdx - 1, 46)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 47), wsOut.Cells(rowIdx - 1, 47)).NumberFormat = "0.00"
    End If

    Call ApplyDarkTheme(wsOut, wsOut.UsedRange, 1)
    wsOut.Columns.AutoFit

    SetStatus wsIn, "寫入其他申報 (8-K/Form 4) 中..."
    Call BuildOtherFilingsSheet(ThisWorkbook, entityName, cik, filings8K, filingsForm4)

    SetStatus wsIn, "繪製 Dashboard（股價圖、財務快照、趨勢圖）中，稍候..."
    Dim priceHistory As Object
    ' Reach back at least as far as the oldest fetched 10-K (wantK years, +1 for
    ' buffer) so Price/Market Cap/P-E/EV-series aren't blank for the oldest
    ' column when B4 asks for more years than the old fixed 6-year window covered.
    Set priceHistory = FetchDailyPrices(ticker, DateAdd("yyyy", -(wantK + 1), Date), Date)
    Call BuildDashboard(ThisWorkbook, ticker, entityName, filings10K, filings10Q, mapEps, priceHistory, _
        mapRevenue, mapShares, allMaps, mapInventory, mapAR, mapCurrentAssets, mapCurrentLiabilities, _
        mapLongTermDebt, mapStockholdersEquity, mapEffectiveTaxRate, mapCapEx, mapCFO, _
        mapCash, mapDA, mapOperatingIncome, mapDividends, mapNetIncome, mapShortTermDebt, wsOut)

    SetStatus wsIn, "完成：" & entityName & "（CIK " & cik & "）10-K " & filings10K.Count & " 筆、10-Q " & filings10Q.Count & " 筆、8-K " & filings8K.Count & " 筆、Form4 " & filingsForm4.Count & " 筆，於 " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    ' B1 still holds whatever SetStatus last wrote before the error, since
    ' this line hasn't overwritten it yet -- surfacing that phase name turns
    ' "型態不符合" into something actually actionable instead of a dead end.
    Dim lastPhase As String
    lastPhase = CStr(wsIn.Range("B1").Value)
    SetStatus wsIn, "發生錯誤（上一步：" & lastPhase & "）：" & Err.Description & " [來源: " & Err.Source & "]"
    Resume CleanExit
End Sub

' Manually-triggered check (button on Input sheet): looks at the company's most
' recent submissions and reports any 10-K/10-Q/8-K/Form 4 filed AFTER the newest
' one already on record for that same form category. Comparing against a max
' filing date (rather than just "accession number not present") matters because
' we deliberately only fetch the top N of each type -- older, intentionally
' un-fetched filings must not be reported as "new". Does not auto-refresh --
' re-enter the ticker in A1 to actually pull the new data in.
Public Sub CheckForNewFilings()
    Dim wsIn As Worksheet
    Set wsIn = ThisWorkbook.Sheets("Input")

    On Error GoTo ErrHandler2

    Dim inputText As String
    inputText = Trim$(CStr(wsIn.Range("A1").Value))
    If inputText = "" Then
        MsgBox "請先在 A1 輸入股票代號並完成一次抓取。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    SetStatus wsIn, "檢查是否有新 filing 中..."

    Dim cik As String, tickerTitle As String, ticker As String
    If Not ResolveCIK(inputText, cik, tickerTitle, ticker) Then
        SetStatus wsIn, "找不到公司：" & inputText
        GoTo CleanExit2
    End If

    Dim json As String
    json = HttpGet("https://data.sec.gov/submissions/CIK" & cik & ".json")
    Dim filingsRaw As String, recentRaw As String
    filingsRaw = ExtractJsonValueRaw(json, "filings")
    recentRaw = ExtractJsonValueRaw(filingsRaw, "recent")

    Dim formsRaw As String, filingDateRaw As String
    formsRaw = ExtractJsonValueRaw(recentRaw, "form")
    filingDateRaw = ExtractJsonValueRaw(recentRaw, "filingDate")

    Dim forms As Collection, fdates As Collection
    Set forms = SplitJsonArrayElements(formsRaw)
    Set fdates = SplitJsonArrayElements(filingDateRaw)

    Dim wsFilings As Worksheet, wsOther As Worksheet
    On Error Resume Next
    Set wsFilings = ThisWorkbook.Sheets("Filings")
    Set wsOther = ThisWorkbook.Sheets("OtherFilings")
    On Error GoTo 0

    Dim max10K As String, max10Q As String, max8K As String, max4 As String
    max10K = MaxDateForForms(wsFilings, 3, 7, "|10-K|10-K/A|")
    max10Q = MaxDateForForms(wsFilings, 3, 7, "|10-Q|10-Q/A|")
    max8K = MaxDateForForms(wsOther, 3, 4, "|8-K|")
    max4 = MaxDateForForms(wsOther, 3, 4, "|4|")

    Dim newList As String
    Dim i As Long
    For i = 1 To forms.Count
        Dim f As String, d As String
        f = JsonUnescape(CStr(forms(i)))
        d = JsonUnescape(CStr(fdates(i)))

        Dim isNew As Boolean
        isNew = False
        Select Case f
            Case "10-K", "10-K/A"
                If max10K = "" Or d > max10K Then isNew = True
            Case "10-Q", "10-Q/A"
                If max10Q = "" Or d > max10Q Then isNew = True
            Case "8-K"
                If max8K = "" Or d > max8K Then isNew = True
            Case "4"
                If max4 = "" Or d > max4 Then isNew = True
        End Select

        If isNew Then
            newList = newList & f & "  (" & d & ")" & vbLf
        End If
    Next i

    If newList = "" Then
        SetStatus wsIn, "已是最新，沒有發現新的 10-K/10-Q/8-K/Form 4（檢查於 " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "）"
    Else
        SetStatus wsIn, "發現尚未抓取的新 filing，請重新輸入 A1 的股票代號以更新"
        MsgBox "發現尚未抓取的新 filing：" & vbLf & vbLf & newList & vbLf & "請重新觸發 A1（例如重新輸入同一個代號並按 Enter）以抓取最新資料。", vbInformation, "有新 filing"
    End If

CleanExit2:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler2:
    SetStatus wsIn, "檢查新 filing 時發生錯誤：" & Err.Description
    Resume CleanExit2
End Sub

' Latest ("yyyy-mm-dd") filingDate already present in ws among rows whose form
' (formCol) is one of matchForms (e.g. "|10-K|10-K/A|"). Returns "" if ws is
' Nothing or no matching row exists.
Private Function MaxDateForForms(ByVal ws As Worksheet, ByVal formCol As Long, ByVal dateCol As Long, ByVal matchForms As String) As String
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim maxD As String
    maxD = ""
    Dim r As Long
    For r = 2 To lastRow
        Dim f As String
        f = CStr(ws.Cells(r, formCol).Value)
        If InStr(matchForms, "|" & f & "|") > 0 Then
            Dim dv As Variant
            dv = ws.Cells(r, dateCol).Value
            Dim d As String
            If IsDate(dv) Then
                d = Format$(dv, "yyyy-mm-dd")
            Else
                d = CStr(dv)
            End If
            If d > maxD Then maxD = d
        End If
    Next r

    MaxDateForForms = maxD
End Function

Private Function GetSetting(ByVal wsIn As Worksheet, ByVal cellAddr As String, ByVal defaultVal As Long) As Long
    Dim v As Variant
    v = wsIn.Range(cellAddr).Value
    If IsNumeric(v) Then
        If CLng(v) > 0 Then
            GetSetting = CLng(v)
            Exit Function
        End If
    End If
    GetSetting = defaultVal
End Function

' ---------- CIK lookup ----------

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

' ---------- Filing list (submissions API) ----------

Private Sub GetFilings(ByVal cik As String, ByVal wantK As Long, ByVal wantQ As Long, ByVal want8K As Long, ByVal want4 As Long, _
    ByRef filings10K As Collection, ByRef filings10Q As Collection, ByRef filings8K As Collection, ByRef filingsForm4 As Collection)

    Set filings10K = New Collection
    Set filings10Q = New Collection
    Set filings8K = New Collection
    Set filingsForm4 = New Collection

    ' Keyed by reportDate so that when a 10-K/A (or 10-Q/A) amends an earlier
    ' original for the same fiscal period, only the amendment is kept. EDGAR's
    ' submissions.json is scanned newest-filed-first (both within the "recent"
    ' block and across the older paginated files below), and an amendment is
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

' ---------- XBRL company facts ----------

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
                            Dim col As New Collection
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

' For flow concepts (revenue, net income, cash flow), 10-Qs commonly tag BOTH the
' standalone quarter AND the year-to-date cumulative figure under the same "end"
' date, distinguished only by "start". We want the quarter-only value, which is
' the one with the LATEST (shortest-period) start date. Balance-sheet concepts
' (Assets, Liabilities) have no "start" at all, so this is a no-op for them.
Public Function LookupConceptValue(ByVal map As Object, ByVal accn As String, ByVal reportDate As String, ByVal targetForm As String) As Variant
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
Public Function LookupFyFp(ByVal maps As Variant, ByVal accn As String, ByVal reportDate As String, ByVal targetForm As String) As String
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

' ---------- Output sheet: Filings (10-K / 10-Q) ----------

' Reads the Accession Number (col 8) / 部門附註連結 (col 24) columns off the
' Filings sheet as it exists BEFORE this run clears it, so WriteFilingRow can
' skip re-querying FindSegmentReportUrl for filings already on record. A
' filing's FilingSummary.xml never changes after the fact, so this cache never
' goes stale -- including a cached empty string ("no segment note found").
Private Function ReadExistingSegmentUrls(ByVal wsOut As Worksheet) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        Set ReadExistingSegmentUrls = d
        Exit Function
    End If

    Dim r As Long
    For r = 2 To lastRow
        Dim accn As String
        accn = CStr(wsOut.Cells(r, 8).Value)
        If accn <> "" And Not d.Exists(accn) Then
            d(accn) = CStr(wsOut.Cells(r, 24).Value)
        End If
    Next r

    Set ReadExistingSegmentUrls = d
End Function

Private Function WriteFilingRow(ByVal wsOut As Worksheet, ByVal rowIdx As Long, ByVal entityName As String, ByVal cik As String, ByVal f As Object, _
    ByVal mapRevenue As Object, ByVal mapNetIncome As Object, ByVal mapEps As Object, ByVal mapAssets As Object, ByVal mapLiabilities As Object, ByVal mapCFO As Object, _
    ByVal mapGrossProfit As Object, ByVal mapOperatingIncome As Object, ByVal mapRD As Object, ByVal mapSGA As Object, ByVal mapDividends As Object, _
    ByVal mapInventory As Object, ByVal mapAR As Object, ByVal mapCurrentAssets As Object, ByVal mapCurrentLiabilities As Object, _
    ByVal mapLongTermDebt As Object, ByVal mapStockholdersEquity As Object, ByVal mapEffectiveTaxRate As Object, ByVal mapCapEx As Object, ByVal mapShares As Object, _
    ByVal allMaps As Variant, ByVal knownSegmentUrls As Object, _
    ByVal mapCOGS As Object, ByVal mapInterestExpense As Object, ByVal mapCash As Object) As Long

    Dim form As String, accn As String, filingDate As String, reportDate As String, primDoc As String, isAmended As Boolean
    form = f("form")
    accn = f("accn")
    filingDate = f("filingDate")
    reportDate = f("reportDate")
    primDoc = f("primaryDocument")
    isAmended = f("amended")

    Dim fyfp As String, parts As Variant
    fyfp = LookupFyFp(allMaps, accn, reportDate, form)
    parts = Split(fyfp, "|")

    Dim revV As Variant, grossV As Variant, opIncV As Variant, assetsV As Variant, liabV As Variant, cfoV As Variant
    revV = LookupConceptValue(mapRevenue, accn, reportDate, form)
    grossV = LookupConceptValue(mapGrossProfit, accn, reportDate, form)
    opIncV = LookupConceptValue(mapOperatingIncome, accn, reportDate, form)
    assetsV = LookupConceptValue(mapAssets, accn, reportDate, form)
    liabV = LookupConceptValue(mapLiabilities, accn, reportDate, form)
    cfoV = LookupConceptValue(mapCFO, accn, reportDate, form)

    Dim curAssetsV As Variant, curLiabV As Variant, equityV As Variant, sharesV As Variant, capexV As Variant
    curAssetsV = LookupConceptValue(mapCurrentAssets, accn, reportDate, form)
    curLiabV = LookupConceptValue(mapCurrentLiabilities, accn, reportDate, form)
    equityV = LookupConceptValue(mapStockholdersEquity, accn, reportDate, form)
    sharesV = LookupConceptValue(mapShares, accn, reportDate, form)
    capexV = LookupConceptValue(mapCapEx, accn, reportDate, form)

    Dim netIncV As Variant, invV As Variant, arV As Variant, cogsV As Variant, intExpV As Variant, cashV As Variant
    netIncV = LookupConceptValue(mapNetIncome, accn, reportDate, form)
    invV = LookupConceptValue(mapInventory, accn, reportDate, form)
    arV = LookupConceptValue(mapAR, accn, reportDate, form)
    cogsV = LookupConceptValue(mapCOGS, accn, reportDate, form)
    intExpV = LookupConceptValue(mapInterestExpense, accn, reportDate, form)
    cashV = LookupConceptValue(mapCash, accn, reportDate, form)

    With wsOut
        .Cells(rowIdx, 1).Value = entityName
        .Cells(rowIdx, 2).Value = cik
        .Cells(rowIdx, 3).Value = form
        .Cells(rowIdx, 4).Value = parts(0)
        .Cells(rowIdx, 5).Value = parts(1)
        .Cells(rowIdx, 6).Value = reportDate
        .Cells(rowIdx, 7).Value = filingDate
        .Cells(rowIdx, 8).Value = accn
        .Cells(rowIdx, 9).Value = BuildDocUrl(cik, accn, primDoc)
        .Cells(rowIdx, 10).Value = SafeNumber(revV)
        .Cells(rowIdx, 11).Value = SafeNumber(netIncV)
        .Cells(rowIdx, 12).Value = SafeNumber(LookupConceptValue(mapEps, accn, reportDate, form))
        .Cells(rowIdx, 13).Value = SafeNumber(assetsV)
        .Cells(rowIdx, 14).Value = SafeNumber(liabV)
        .Cells(rowIdx, 15).Value = SafeNumber(cfoV)
        .Cells(rowIdx, 16).Value = IIf(isAmended, "是", "")
        .Cells(rowIdx, 17).Value = SafeRatio(grossV, revV)
        .Cells(rowIdx, 18).Value = SafeRatio(opIncV, revV)
        .Cells(rowIdx, 19).Value = SafeNumber(LookupConceptValue(mapRD, accn, reportDate, form))
        .Cells(rowIdx, 20).Value = SafeNumber(LookupConceptValue(mapSGA, accn, reportDate, form))
        .Cells(rowIdx, 21).Value = SafeRatio(liabV, assetsV)
        .Cells(rowIdx, 22).Value = SafeNumber(LookupConceptValue(mapDividends, accn, reportDate, form))
        .Cells(rowIdx, 23).Value = BuildIndexUrl(cik, accn)
        If Not knownSegmentUrls Is Nothing And knownSegmentUrls.Exists(accn) Then
            .Cells(rowIdx, 24).Value = knownSegmentUrls(accn)
        Else
            .Cells(rowIdx, 24).Value = FindSegmentReportUrl(cik, accn)
        End If
        .Cells(rowIdx, 25).Value = SafeNumber(invV)
        .Cells(rowIdx, 26).Value = SafeNumber(arV)
        .Cells(rowIdx, 27).Value = SafeNumber(curAssetsV)
        .Cells(rowIdx, 28).Value = SafeNumber(curLiabV)
        .Cells(rowIdx, 29).Value = SafeRatio(curAssetsV, curLiabV)
        .Cells(rowIdx, 30).Value = SafeNumber(LookupConceptValue(mapLongTermDebt, accn, reportDate, form))
        .Cells(rowIdx, 31).Value = SafeNumber(equityV)
        .Cells(rowIdx, 32).Value = SafeRatio(equityV, sharesV)
        .Cells(rowIdx, 33).Value = SafeNumber(LookupConceptValue(mapEffectiveTaxRate, accn, reportDate, form))
        .Cells(rowIdx, 34).Value = SafeNumber(capexV)
        .Cells(rowIdx, 35).Value = SafeSubtract(cfoV, capexV)
        .Cells(rowIdx, 36).Value = SafeRatio(netIncV, revV)
        .Cells(rowIdx, 37).Value = SafeRatio(netIncV, equityV)
        .Cells(rowIdx, 38).Value = SafeRatio(netIncV, assetsV)
        .Cells(rowIdx, 39).Value = SafeRatio(cogsV, invV)
        .Cells(rowIdx, 40).Value = SafeRatio(revV, arV)
        .Cells(rowIdx, 41).Value = SafeRatio(revV, assetsV)
        .Cells(rowIdx, 42).Value = SafeRatio(revV, cashV)
        .Cells(rowIdx, 43).Value = SafeRatio(SafeSubtract(curAssetsV, invV), curLiabV)
        .Cells(rowIdx, 44).Value = SafeRatio(liabV, equityV)
        .Cells(rowIdx, 45).Value = SafeRatio(assetsV, equityV)
        .Cells(rowIdx, 46).Value = SafeNumber(intExpV)
        .Cells(rowIdx, 47).Value = SafeRatio(opIncV, intExpV)
    End With

    WriteFilingRow = rowIdx + 1
End Function

' ---------- Output sheet: OtherFilings (8-K / Form 4) ----------

Public Sub BuildOtherFilingsSheet(ByVal wb As Workbook, ByVal entityName As String, ByVal cik As String, ByVal filings8K As Collection, ByVal filingsForm4 As Collection)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("OtherFilings")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "OtherFilings"
    Else
        ws.Cells.Clear
    End If

    Dim headers As Variant
    headers = Array("公司名稱", "CIK", "表單類別", "申報日期", "事件代碼/說明", "文件連結", "申報索引頁", "Accession Number")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

    Dim rowIdx As Long
    rowIdx = 2
    Dim f As Variant
    For Each f In filings8K
        rowIdx = WriteOtherFilingRow(ws, rowIdx, entityName, cik, f)
    Next f
    For Each f In filingsForm4
        rowIdx = WriteOtherFilingRow(ws, rowIdx, entityName, cik, f)
    Next f

    If rowIdx > 2 Then
        ws.Range(ws.Cells(2, 4), ws.Cells(rowIdx - 1, 4)).NumberFormat = "yyyy-mm-dd"
    End If

    Call ApplyDarkTheme(ws, ws.UsedRange, 1)
    ws.Columns.AutoFit
End Sub

Private Function WriteOtherFilingRow(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal entityName As String, ByVal cik As String, ByVal f As Object) As Long
    Dim form As String, accn As String, filingDate As String, primDoc As String, itemsStr As String
    form = f("form")
    accn = f("accn")
    filingDate = f("filingDate")
    primDoc = f("primaryDocument")
    itemsStr = f("items")

    With ws
        .Cells(rowIdx, 1).Value = entityName
        .Cells(rowIdx, 2).Value = cik
        .Cells(rowIdx, 3).Value = form
        .Cells(rowIdx, 4).Value = filingDate
        .Cells(rowIdx, 5).Value = itemsStr
        .Cells(rowIdx, 6).Value = BuildDocUrl(cik, accn, primDoc)
        .Cells(rowIdx, 7).Value = BuildIndexUrl(cik, accn)
        .Cells(rowIdx, 8).Value = accn
    End With
    WriteOtherFilingRow = rowIdx + 1
End Function

' ---------- URL builders ----------

Private Function BuildDocUrl(ByVal cik As String, ByVal accn As String, ByVal primaryDocument As String) As String
    Dim cikNoLead As String
    cikNoLead = CStr(CLng(cik))
    Dim accnNoDash As String
    accnNoDash = Replace(accn, "-", "")
    If primaryDocument <> "" Then
        BuildDocUrl = "https://www.sec.gov/Archives/edgar/data/" & cikNoLead & "/" & accnNoDash & "/" & primaryDocument
    Else
        BuildDocUrl = "https://www.sec.gov/Archives/edgar/data/" & cikNoLead & "/" & accn & ".txt"
    End If
End Function

Private Function BuildIndexUrl(ByVal cik As String, ByVal accn As String) As String
    Dim cikNoLead As String, accnNoDash As String
    cikNoLead = CStr(CLng(cik))
    accnNoDash = Replace(accn, "-", "")
    BuildIndexUrl = "https://www.sec.gov/Archives/edgar/data/" & cikNoLead & "/" & accnNoDash & "/" & accn & "-index.htm"
End Function

' The companyfacts/companyconcept XBRL APIs only expose non-dimensional (i.e.
' consolidated total) facts -- segment/geographic breakdowns only exist in the
' filing's raw XBRL instance document, which is too filer-specific to parse
' reliably here. This instead points at the auto-generated "Segment" note page
' (if the filing's Financial Report viewer has one), found via FilingSummary.xml.
Private Function FindSegmentReportUrl(ByVal cik As String, ByVal accn As String) As String
    Dim cikNoLead As String, accnNoDash As String
    cikNoLead = CStr(CLng(cik))
    accnNoDash = Replace(accn, "-", "")

    Dim url As String
    url = "https://www.sec.gov/Archives/edgar/data/" & cikNoLead & "/" & accnNoDash & "/FilingSummary.xml"

    Dim xml As String
    On Error Resume Next
    xml = HttpGet(url)
    On Error GoTo 0
    If Len(xml) = 0 Then
        FindSegmentReportUrl = ""
        Exit Function
    End If

    Dim chunks() As String
    chunks = Split(xml, "<Report ")
    Dim i As Long
    For i = 1 To UBound(chunks)
        Dim chunk As String
        chunk = chunks(i)
        Dim shortName As String, longName As String, htmlFile As String
        shortName = ExtractXmlTag(chunk, "ShortName")
        longName = ExtractXmlTag(chunk, "LongName")
        htmlFile = ExtractXmlTag(chunk, "HtmlFileName")
        If htmlFile <> "" Then
            If InStr(1, shortName, "segment", vbTextCompare) > 0 Or InStr(1, longName, "segment", vbTextCompare) > 0 Then
                FindSegmentReportUrl = "https://www.sec.gov/Archives/edgar/data/" & cikNoLead & "/" & accnNoDash & "/" & htmlFile
                Exit Function
            End If
        End If
    Next i

    FindSegmentReportUrl = ""
End Function

Private Function ExtractXmlTag(ByVal s As String, ByVal tagName As String) As String
    Dim startTag As String, endTag As String
    startTag = "<" & tagName & ">"
    endTag = "</" & tagName & ">"
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, s, startTag, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(startTag)
    p2 = InStr(p1, s, endTag, vbTextCompare)
    If p2 = 0 Then Exit Function
    ExtractXmlTag = Mid$(s, p1, p2 - p1)
End Function

' ---------- Numeric helpers ----------

Private Function SafeNumber(ByVal v As Variant) As Variant
    If IsNull(v) Then
        SafeNumber = ""
    ElseIf CStr(v) = "" Then
        SafeNumber = ""
    ElseIf IsNumeric(v) Then
        SafeNumber = CDbl(v)
    Else
        SafeNumber = v
    End If
End Function

Private Function SafeRatio(ByVal numerator As Variant, ByVal denominator As Variant) As Variant
    If IsNumeric(numerator) And CStr(numerator) <> "" And IsNumeric(denominator) And CStr(denominator) <> "" Then
        If CDbl(denominator) <> 0 Then
            SafeRatio = CDbl(numerator) / CDbl(denominator)
            Exit Function
        End If
    End If
    SafeRatio = ""
End Function

Private Function SafeSubtract(ByVal a As Variant, ByVal b As Variant) As Variant
    If IsNumeric(a) And CStr(a) <> "" And IsNumeric(b) And CStr(b) <> "" Then
        SafeSubtract = CDbl(a) - CDbl(b)
    Else
        SafeSubtract = ""
    End If
End Function

' ---------- Sheet / status helpers ----------

Private Sub ClearOutputSheet(ByVal wsOut As Worksheet)
    wsOut.Cells.Clear
End Sub

Private Sub WriteHeaders(ByVal wsOut As Worksheet)
    Dim headers As Variant
    headers = Array("公司名稱", "CIK", "表單類別", "會計年度(FY)", "會計期間(FP)", "財報截止日", "申報日期", "Accession Number", "文件連結", _
        "營收 Revenue", "淨利 Net Income", "稀釋EPS", "總資產 Assets", "總負債 Liabilities", "營運現金流 CFO", _
        "修正版", "毛利率 Gross Margin", "營業利益率 Operating Margin", "研發費用 R&D", "SG&A", "負債比 Debt Ratio", "每股股利 Dividend/Share", _
        "申報索引頁", "部門附註連結(如有)", _
        "存貨 Inventory", "應收帳款 AR", "流動資產 Current Assets", "流動負債 Current Liabilities", "流動比率 Current Ratio", _
        "長期負債 LT Debt", "股東權益 Equity", "每股淨值 Book Value/Share", "有效稅率 Tax Rate", "資本支出 CapEx", "自由現金流 FCF", _
        "淨利率 Net Margin", "股東權益報酬率 ROE", "資產報酬率 ROA", "存貨週轉率 Inventory Turnover", "應收帳款週轉率 AR Turnover", _
        "總資產週轉率 Asset Turnover", "現金週轉率 Cash Turnover", "速動比率 Quick Ratio", "負債權益比 Debt/Equity", "權益乘數 Equity Multiplier", _
        "利息費用 Interest Expense", "利息保障倍數 Interest Coverage")
    Dim i As Long
    For i = 0 To UBound(headers)
        wsOut.Cells(1, i + 1).Value = headers(i)
    Next i
    ' actual header styling (black/orange, Calibri) is applied afterward by
    ' ApplyDarkTheme once all data rows are written
End Sub

Private Sub SetStatus(ByVal wsIn As Worksheet, ByVal msg As String)
    wsIn.Range("B1").Value = msg
    Application.StatusBar = msg
End Sub

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

' ---------- Watchlist (peer / multi-ticker comparison) ----------

' Triggered by the Watchlist sheet's own Worksheet_Change (see
' SheetWatchlist_Code.txt) when a single ticker cell in the list changes.
' Deliberately lightweight compared to FetchSECFilings: only the latest 10-K's
' figures are pulled (no multi-year history, no per-filing document links, no
' charts), and only the one changed row is touched -- entering ticker #16 in a
' 15-ticker watchlist doesn't re-fetch the other 15.
Public Sub RefreshWatchlistRow(ByVal ws As Worksheet, ByVal r As Long, ByVal tickerInput As String)
    On Error GoTo ErrHandler3
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Cells(r, 2).Value = "更新中..."
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 11)).ClearContents

    Dim cik As String, tickerTitle As String, ticker As String
    If Not ResolveCIK(tickerInput, cik, tickerTitle, ticker) Then
        ws.Cells(r, 2).Value = "找不到公司"
        ws.Cells(r, 12).Value = Format$(Now, "hh:nn:ss")
        GoTo CleanExit3
    End If

    Dim filings10K As Collection, filings10Q As Collection, filings8K As Collection, filingsForm4 As Collection
    Call GetFilings(cik, 1, 0, 0, 0, filings10K, filings10Q, filings8K, filingsForm4)
    If filings10K.Count = 0 Then
        ws.Cells(r, 2).Value = "無 10-K 資料"
        ws.Cells(r, 12).Value = Format$(Now, "hh:nn:ss")
        GoTo CleanExit3
    End If

    Dim f As Object
    Set f = filings10K(1)
    Dim accn As String, reportDate As String, form As String
    accn = f("accn"): reportDate = f("reportDate"): form = f("form")

    Dim companyFactsJson As String
    companyFactsJson = HttpGet("https://data.sec.gov/api/xbrl/companyfacts/CIK" & cik & ".json")
    Dim entityName As String
    entityName = ExtractJsonString(companyFactsJson, "entityName")
    If entityName = "" Then entityName = tickerTitle

    Dim factsRaw As String, usgaapRaw As String
    factsRaw = ExtractJsonValueRaw(companyFactsJson, "facts")
    usgaapRaw = ExtractJsonValueRaw(factsRaw, "us-gaap")

    Dim mapRevenue As Object, mapNetIncome As Object, mapEps As Object, mapShares As Object
    Dim mapStockholdersEquity As Object, mapLongTermDebt As Object, mapShortTermDebt As Object
    Dim mapCash As Object, mapOperatingIncome As Object, mapDA As Object, mapDividends As Object
    Dim mapCurrentAssets As Object, mapCurrentLiabilities As Object

    Set mapRevenue = BuildConceptMap(usgaapRaw, Array("Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"), "USD")
    Set mapNetIncome = BuildConceptMap(usgaapRaw, Array("NetIncomeLoss", "ProfitLoss"), "USD")
    Set mapEps = BuildConceptMap(usgaapRaw, Array("EarningsPerShareDiluted", "EarningsPerShareBasicAndDiluted"), "USD/shares")
    Set mapShares = BuildConceptMap(usgaapRaw, Array("CommonStockSharesOutstanding"), "shares")
    Set mapStockholdersEquity = BuildConceptMap(usgaapRaw, Array("StockholdersEquity"), "USD")
    Set mapLongTermDebt = BuildConceptMap(usgaapRaw, Array("LongTermDebtNoncurrent", "LongTermDebt"), "USD")
    Set mapShortTermDebt = BuildConceptMap(usgaapRaw, Array("ShortTermBorrowings", "DebtCurrent", "LongTermDebtCurrent"), "USD")
    Set mapCash = BuildConceptMap(usgaapRaw, Array("CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"), "USD")
    Set mapOperatingIncome = BuildConceptMap(usgaapRaw, Array("OperatingIncomeLoss"), "USD")
    Set mapDA = BuildConceptMap(usgaapRaw, Array("DepreciationDepletionAndAmortization", "DepreciationAmortizationAndAccretionNet", "DepreciationAndAmortization"), "USD")
    Set mapDividends = BuildConceptMap(usgaapRaw, Array("CommonStockDividendsPerShareDeclared", "CommonStockDividendsPerShareCashPaid"), "USD/shares")
    Set mapCurrentAssets = BuildConceptMap(usgaapRaw, Array("AssetsCurrent"), "USD")
    Set mapCurrentLiabilities = BuildConceptMap(usgaapRaw, Array("LiabilitiesCurrent"), "USD")

    Dim revV As Variant, netIncV As Variant, epsV As Variant, sharesV As Variant
    Dim equityV As Variant, ltDebtV As Variant, stDebtV As Variant, cashV As Variant
    Dim opIncV As Variant, daV As Variant, dpsV As Variant, curAssetsV As Variant, curLiabV As Variant
    revV = LookupConceptValue(mapRevenue, accn, reportDate, form)
    netIncV = LookupConceptValue(mapNetIncome, accn, reportDate, form)
    epsV = LookupConceptValue(mapEps, accn, reportDate, form)
    sharesV = LookupConceptValue(mapShares, accn, reportDate, form)
    equityV = LookupConceptValue(mapStockholdersEquity, accn, reportDate, form)
    ltDebtV = LookupConceptValue(mapLongTermDebt, accn, reportDate, form)
    stDebtV = LookupConceptValue(mapShortTermDebt, accn, reportDate, form)
    cashV = LookupConceptValue(mapCash, accn, reportDate, form)
    opIncV = LookupConceptValue(mapOperatingIncome, accn, reportDate, form)
    daV = LookupConceptValue(mapDA, accn, reportDate, form)
    dpsV = LookupConceptValue(mapDividends, accn, reportDate, form)
    curAssetsV = LookupConceptValue(mapCurrentAssets, accn, reportDate, form)
    curLiabV = LookupConceptValue(mapCurrentLiabilities, accn, reportDate, form)

    Dim revGrowth As Variant
    revGrowth = PriorYearGrowth(mapRevenue, accn, reportDate)

    Dim priceV As Variant
    priceV = LatestPrice(ticker)

    Dim mktCapV As Variant, evV As Variant, ebitdaV As Variant
    mktCapV = SafeMultiply(priceV, sharesV)

    Dim debtForEV As Double
    debtForEV = 0
    If IsNumeric(ltDebtV) And CStr(ltDebtV) <> "" Then debtForEV = debtForEV + CDbl(ltDebtV)
    If IsNumeric(stDebtV) And CStr(stDebtV) <> "" Then debtForEV = debtForEV + CDbl(stDebtV)
    If IsNumeric(mktCapV) And CStr(mktCapV) <> "" And IsNumeric(cashV) And CStr(cashV) <> "" Then
        evV = CDbl(mktCapV) + debtForEV - CDbl(cashV)
    Else
        evV = ""
    End If

    If IsNumeric(opIncV) And CStr(opIncV) <> "" And IsNumeric(daV) And CStr(daV) <> "" Then
        ebitdaV = CDbl(opIncV) + CDbl(daV)
    Else
        ebitdaV = ""
    End If

    With ws
        .Cells(r, 1).Value = ticker
        .Cells(r, 2).Value = entityName
        .Cells(r, 3).Value = SafeNumber(priceV)
        .Cells(r, 4).Value = SafeNumber(mktCapV)
        .Cells(r, 5).Value = revGrowth
        .Cells(r, 6).Value = SafeRatio(priceV, epsV)
        .Cells(r, 7).Value = SafeRatio(mktCapV, revV)
        .Cells(r, 8).Value = SafeRatio(evV, ebitdaV)
        .Cells(r, 9).Value = SafeRatio(netIncV, equityV)
        .Cells(r, 10).Value = SafeRatio(dpsV, priceV)
        .Cells(r, 11).Value = SafeRatio(curAssetsV, curLiabV)
        .Cells(r, 12).Value = Format$(Now, "hh:nn:ss")
    End With

    ws.Cells(r, 3).NumberFormat = "#,##0.00"
    ws.Cells(r, 4).NumberFormat = "#,##0,,""M"""
    ws.Cells(r, 5).NumberFormat = "0.00%"
    ws.Range(ws.Cells(r, 6), ws.Cells(r, 8)).NumberFormat = "0.00"
    ws.Cells(r, 9).NumberFormat = "0.00%"
    ws.Cells(r, 10).NumberFormat = "0.00%"
    ws.Cells(r, 11).NumberFormat = "0.00"

CleanExit3:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler3:
    ws.Cells(r, 2).Value = "錯誤：" & Err.Description
    ws.Cells(r, 12).Value = Format$(Now, "hh:nn:ss")
    Resume CleanExit3
End Sub

' Clears a Watchlist row's fetched columns (B:L) when the user deletes the
' ticker in column A, so a blanked-out ticker doesn't leave stale data behind.
Public Sub ClearWatchlistRow(ByVal ws As Worksheet, ByVal r As Long)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 12)).ClearContents
End Sub

' A 10-K's XBRL facts report multiple fiscal years of comparatives under the
' SAME accession number (accn) -- e.g. the current year (end=reportDate) plus
' 1-2 prior years' income-statement figures, all filed together. This finds
' the entry roughly one year before reportDate (300-400 days, allowing for
' 52/53-week fiscal calendars shifting the exact date) within that same accn
' and returns (current/prior - 1), or "" if either side is missing.
Private Function PriorYearGrowth(ByVal map As Object, ByVal accn As String, ByVal reportDate As String) As Variant
    PriorYearGrowth = ""
    If map Is Nothing Then Exit Function
    If Not map.Exists(accn) Then Exit Function
    If Not IsDate(reportDate) Then Exit Function

    Dim col As Collection
    Set col = map(accn)

    Dim curVal As Double, curFound As Boolean
    Dim bestPriorVal As Double, bestPriorDiff As Long, priorFound As Boolean
    curFound = False
    priorFound = False

    Dim e As Variant
    For Each e In col
        Dim endD As String, valRaw As String
        endD = ExtractJsonString(CStr(e), "end")
        If endD = reportDate Then
            valRaw = ExtractJsonValueRaw(CStr(e), "val")
            If IsNumeric(valRaw) Then
                curVal = CDbl(valRaw)
                curFound = True
            End If
        ElseIf IsDate(endD) Then
            Dim diffDays As Long
            diffDays = DateDiff("d", CDate(endD), CDate(reportDate))
            If diffDays >= 300 And diffDays <= 400 Then
                valRaw = ExtractJsonValueRaw(CStr(e), "val")
                If IsNumeric(valRaw) Then
                    If Not priorFound Or Abs(diffDays - 365) < bestPriorDiff Then
                        bestPriorVal = CDbl(valRaw)
                        bestPriorDiff = Abs(diffDays - 365)
                        priorFound = True
                    End If
                End If
            End If
        End If
    Next e

    If curFound And priorFound And bestPriorVal <> 0 Then
        PriorYearGrowth = curVal / bestPriorVal - 1
    End If
End Function

Private Function SafeMultiply(ByVal a As Variant, ByVal b As Variant) As Variant
    If IsNumeric(a) And CStr(a) <> "" And IsNumeric(b) And CStr(b) <> "" Then
        SafeMultiply = CDbl(a) * CDbl(b)
    Else
        SafeMultiply = ""
    End If
End Function
