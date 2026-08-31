Attribute VB_Name = "modMOPS"
Option Explicit

' Taiwan MOPS (mopsov.twse.com.tw) financial-statement fetcher, built to land TW
' tickers on the SAME Dashboard/RawData/chart machinery as US tickers (see
' modCharts.BuildDashboard's dashboardSheetName/rawDataSheetName/annualFormMarker
' params, added for this purpose) instead of a parallel implementation.
'
' MOPS's t164sb01 report is INLINE XBRL: every reported number sits inside a tag
' like <ix:nonFraction name="ifrs-full:Revenue" contextRef="From20260401To20260630"
' scale="3" sign="-"?>49,120,964</ix:nonFraction>, tagged with the standard
' ifrs-full taxonomy (plus filer-extension tifrs-* tags for a handful of concepts
' the base taxonomy has no slot for, e.g. AR net / COGS). This is what this module
' parses -- NOT the rendered <td> table text, which is what a human (or WebFetch)
' reads.
'
' Two important differences from SEC's XBRL that shaped the design below:
'   1. SEC's companyfacts.json is ONE call covering every historical period ever
'      filed; MOPS only ever shows ONE quarter (plus its prior-year comparative)
'      per HTTP request. Building an N-quarter trend means N separate fetches,
'      walked backward from today (see FetchMOPSFilings's quarter-decrementing
'      loop) -- there is no equivalent of GetFilings's single submissions.json call.
'   2. A MOPS quarterly report's income-statement/cash-flow contexts expose BOTH
'      the standalone quarter (From<qStart>To<qEnd>) AND the cumulative
'      fiscal-year-to-date figure (From<fyStart>To<qEnd>) as separate, explicit
'      contexts -- unlike SEC filings, there is no ambiguity to resolve via
'      "shortest start date wins": this module reads the exact contextRef it
'      wants directly (see ExtractTWMetrics's durCtx).
'
' Known limitation: iXBRL tagging on this endpoint was verified working for
' 3017's FY2026Q2 report; how far back it extends before filings stop being
' inline-XBRL-tagged has not been verified. A quarter that yields zero
' <ix:nonFraction> matches is simply skipped (see the facts.Count = 0 check),
' not treated as an error -- older quarters may quietly drop out of the trend
' rather than fail the whole run.

Private Const MOPS_USER_AGENT As String = "Ryan Personal Research Tool ryan929929@gmail.com"

' Every metric bucket this module extracts per period, shared between the
' concept-map accumulation (feeds BuildDashboard's snapshot table) and the
' TW_Filings row writer (feeds BuildFinancialChartsInto's trend charts).
Private Function MetricKeys() As Variant
    MetricKeys = Array("Revenue", "GrossProfit", "OperatingIncome", "NetIncome", "Eps", "Assets", "Liabilities", "Cfo", _
        "Rd", "Sga", "Inventory", "Ar", "CurrentAssets", "CurrentLiabilities", "LongTermDebt", "ShortTermDebt", _
        "StockholdersEquity", "CapEx", "Cash", "Da", "Cogs", "InterestExpense", "EffectiveTaxRate", "Shares", "AccountsPayable")
End Function

' ---------- Entry point ----------

' Triggered by the Input sheet's Worksheet_Change when the TW ticker cell (B9)
' changes. Mirrors modSEC.FetchSECFilings's shape: resolve settings, fetch,
' write Filings-equivalent sheet, build Dashboard, build QuarterlySnapshot.
Public Sub FetchMOPSFilings()
    Dim wsIn As Worksheet, wsOut As Worksheet
    Set wsIn = ThisWorkbook.Sheets("Input")
    Set wsOut = modCharts.GetOrCreateSheet(ThisWorkbook, "TW_Filings")

    On Error GoTo ErrHandler

    Dim coId As String
    coId = Trim$(CStr(wsIn.Range("B9").Value))
    If coId = "" Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    SetTWStatus wsIn, "查詢台股 " & coId & " 財報中..."

    Dim wantQ As Long
    wantQ = GetTWSetting(wsIn, "B10", 10)

    ' Back off 95 days from today before starting the walk so the very first
    ' quarter tried is essentially always already published -- Q1-Q3 reviewed
    ' reports land ~45 days after quarter-end, the Q4/annual audited report
    ' takes ~75-90 days, so 95 clears both with room to spare.
    Dim curYear As Long, curSeason As Long
    Call QuarterOf(DateAdd("d", -95, Date), curYear, curSeason)

    Dim filingsAnnual As New Collection, filingsQuarterly As New Collection
    Dim annualRows As New Collection, quarterlyRows As New Collection
    ' ONE shared map set, not separate annual/quarterly ones -- BuildDashboard's
    ' internal price chart (BuildPriceChartInto) walks BOTH filings10K and
    ' filings10Q against a SINGLE mapEps, so every metric's accn->value map must
    ' contain both the annual (…FY) and quarterly (…Q#) accessions side by side,
    ' exactly like SEC's companyfacts-derived maps already do for 10-K/10-Q accns.
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

        SetTWStatus wsIn, "抓取 " & curYear & "Q" & curSeason & " 財報中 (" & (gotQ + 1) & "/" & wantQ & ")..."

        Dim usedUrl As String, fetchedName As String
        Dim facts As Object
        Set facts = FetchQuarterFacts(coId, curYear, curSeason, usedUrl, fetchedName)

        If facts.Count > 0 Then
            If entityName = "" Then entityName = fetchedName

            Dim qStart As Date, qEnd As Date, fyStart As Date
            Call QuarterBounds(curYear, curSeason, qStart, qEnd, fyStart)

            Dim accnQ As String, fyLabel As String, fpLabelQ As String
            fyLabel = CStr(curYear)
            fpLabelQ = "Q" & curSeason
            accnQ = "TW" & coId & "-" & fyLabel & fpLabelQ

            Dim metricsQ As Object
            Set metricsQ = ExtractTWMetrics(facts, qStart, qEnd, qEnd)
            Call AccumulateMetricsIntoMaps(maps, metricsQ, accnQ, "MOPS-Q", fyLabel, fpLabelQ, qStart, qEnd, qEnd)
            filingsQuarterly.Add BuildTWFilingDict(accnQ, "MOPS-Q", qEnd)
            quarterlyRows.Add BuildTWRowSpec(entityName, coId, metricsQ, "MOPS-Q", fyLabel, fpLabelQ, qEnd, usedUrl)
            gotQ = gotQ + 1

            If curSeason = 4 Then
                Dim accnA As String, fpLabelA As String
                fpLabelA = "FY"
                accnA = "TW" & coId & "-" & fyLabel & fpLabelA

                Dim metricsA As Object
                Set metricsA = ExtractTWMetrics(facts, fyStart, qEnd, qEnd)
                Call AccumulateMetricsIntoMaps(maps, metricsA, accnA, "MOPS-10-K", fyLabel, fpLabelA, fyStart, qEnd, qEnd)
                filingsAnnual.Add BuildTWFilingDict(accnA, "MOPS-10-K", qEnd)
                annualRows.Add BuildTWRowSpec(entityName, coId, metricsA, "MOPS-10-K", fyLabel, fpLabelA, qEnd, usedUrl)
            End If
        End If

        curSeason = curSeason - 1
        If curSeason = 0 Then
            curSeason = 4
            curYear = curYear - 1
        End If
    Loop

    If filingsQuarterly.Count = 0 Then
        SetTWStatus wsIn, "找不到台股 " & coId & " 的 MOPS 財報資料，請確認代號"
        GoTo CleanExit
    End If
    If entityName = "" Then entityName = coId

    SetTWStatus wsIn, "寫入 TW_Filings 工作表中..."
    Call ClearTWFilingsSheet(wsOut)
    Call WriteTWHeaders(wsOut)

    Dim rowIdx As Long
    rowIdx = 2
    Dim spec As Variant
    For Each spec In annualRows
        Call WriteTWFilingRow(wsOut, rowIdx, spec)
        rowIdx = rowIdx + 1
    Next spec
    For Each spec In quarterlyRows
        Call WriteTWFilingRow(wsOut, rowIdx, spec)
        rowIdx = rowIdx + 1
    Next spec

    If rowIdx > 2 Then
        wsOut.Range(wsOut.Cells(2, 6), wsOut.Cells(rowIdx - 1, 7)).NumberFormat = "yyyy-mm-dd"
        wsOut.Range(wsOut.Cells(2, 10), wsOut.Cells(rowIdx - 1, 11)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 12), wsOut.Cells(rowIdx - 1, 12)).NumberFormat = "0.00"
        wsOut.Range(wsOut.Cells(2, 13), wsOut.Cells(rowIdx - 1, 15)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 17), wsOut.Cells(rowIdx - 1, 18)).NumberFormat = "0.00%"
        wsOut.Range(wsOut.Cells(2, 19), wsOut.Cells(rowIdx - 1, 20)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(2, 21), wsOut.Cells(rowIdx - 1, 21)).NumberFormat = "0.00%"
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

    Call modTheme.ApplyDarkTheme(wsOut, wsOut.UsedRange, 1)
    wsOut.Columns.AutoFit

    Dim filingsBlockLastRow As Long
    filingsBlockLastRow = rowIdx - 1

    SetTWStatus wsIn, "繪製台股 Dashboard 中，稍候..."

    ' Any metric map works here as long as it's populated for essentially every
    ' accn (used only as a fy/fp lookup fallback chain by LookupFyFp) -- same
    ' array serves both calls below since both draw accns from the same "maps".
    Dim allMapsArr As Variant
    allMapsArr = Array(maps("Assets"), maps("Revenue"), maps("NetIncome"))

    Dim priceHistory As Object
    Set priceHistory = FetchTWPriceHistory(coId, wantQ)

    Call modCharts.BuildDashboard(ThisWorkbook, coId, entityName, filingsAnnual, filingsQuarterly, maps("Eps"), priceHistory, _
        maps("Revenue"), maps("Shares"), allMapsArr, maps("Inventory"), maps("Ar"), maps("CurrentAssets"), maps("CurrentLiabilities"), _
        maps("LongTermDebt"), maps("StockholdersEquity"), maps("EffectiveTaxRate"), maps("CapEx"), maps("Cfo"), _
        maps("Cash"), maps("Da"), maps("OperatingIncome"), maps("Cfo"), maps("NetIncome"), _
        maps("ShortTermDebt"), maps("Cogs"), maps("AccountsPayable"), maps("Assets"), filingsBlockLastRow, wsOut, "TW_Dashboard", "TW_RawData", "10-K")
    ' Note: maps("Cfo") is passed twice deliberately -- see the "dividends"
    ' comment in ExtractTWMetrics: there is no reliable per-share dividend tag
    ' in this feed, so the Dividends slot has nothing real to pass. Reusing Cfo
    ' here (rather than Nothing) avoids a LookupConceptValue crash on a Nothing
    ' map; DPS/Dividend Yield/Payout Ratio on the resulting Dashboard read as
    ' CFO-derived nonsense and are overwritten to blank below.
    Call BlankOutDividendRow(ThisWorkbook, "TW_Dashboard")

    SetTWStatus wsIn, "繪製台股季度快照分頁中..."
    Call modCharts.BuildQuarterlyDashboard(ThisWorkbook, coId, entityName, filingsQuarterly, maps("Revenue"), maps("Eps"), maps("Shares"), allMapsArr, priceHistory, _
        maps("Inventory"), maps("Ar"), maps("CurrentAssets"), maps("CurrentLiabilities"), _
        maps("LongTermDebt"), maps("StockholdersEquity"), maps("EffectiveTaxRate"), maps("CapEx"), maps("Cfo"), _
        maps("Cash"), maps("Da"), maps("OperatingIncome"), maps("Cfo"), maps("NetIncome"), _
        maps("ShortTermDebt"), maps("Cogs"), maps("AccountsPayable"), maps("Assets"), "TW_QuarterlySnapshot")
    ' Same CFO-placeholder-for-dividends issue as the annual Dashboard call
    ' above -- BuildQuarterlyDashboard was ALSO given maps("Cfo") in the
    ' dividends slot, so TW_QuarterlySnapshot needs the same cleanup call
    ' (confirmed missing: without this, its Dividend Yield/Payout Ratio rows
    ' showed CFO-derived nonsense in the millions of percent).
    Call BlankOutDividendRow(ThisWorkbook, "TW_QuarterlySnapshot")

    SetTWStatus wsIn, "抓取月營收中..."
    Call BuildMonthlyRevenue(wsIn, coId, entityName, wantQ * 3 + 6)

    SetTWStatus wsIn, "匯出個股快照 (.xlsx) 中..."
    Call modCharts.ExportTickerSnapshot(ThisWorkbook, coId, "TW_Dashboard", "TW_RawData", "TW_QuarterlySnapshot")

    SetTWStatus wsIn, "完成：" & entityName & "（" & coId & "）季報 " & filingsQuarterly.Count & " 筆、年報 " & filingsAnnual.Count & " 筆，於 " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Dim lastPhase As String
    lastPhase = CStr(wsIn.Range("B1").Value)
    SetTWStatus wsIn, "台股抓取發生錯誤（上一步：" & lastPhase & "）：" & Err.Description & " [來源: " & Err.Source & "]"
    Resume CleanExit
End Sub

' Dividends per share has no reliable direct tag in this feed (see
' ExtractTWMetrics's comment) -- FetchMOPSFilings passes the CFO map into
' BuildDashboard's dividends slot purely so LookupConceptValue always has a
' real Dictionary to call .Exists on, then this blanks the resulting nonsense
' cells (Dividend Yield / Payout Ratio rows, and the DPS row) back out rather
' than showing CFO-derived garbage.
Private Sub BlankOutDividendRow(ByVal wb As Workbook, ByVal dashboardSheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(dashboardSheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then Exit Sub

    Dim r As Long
    For r = 1 To lastRow
        Dim lbl As String
        lbl = CStr(ws.Cells(r, 1).Value)
        If lbl = "Dividend Yield" Or lbl = "Payout Ratio" Then
            ws.Range(ws.Cells(r, 2), ws.Cells(r, lastCol)).ClearContents
        End If
    Next r
End Sub

' ---------- Per-quarter fetch + parse ----------

' Fetches one (CO_ID, SYEAR, SSEASON) quarterly report, trying REPORT_ID=C
' (consolidated) first and falling back to A (individual/parent-only, for
' filers with no subsidiaries to consolidate) when C comes back with zero
' taggedfacts. usedUrlOut reports whichever URL actually worked, for the
' Filings sheet's "文件連結"/"申報索引頁" columns. entityNameOut is extracted
' here (rather than making the caller re-fetch/re-parse the HTML) since this
' is the only place the raw HTML string is in scope.
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

' ---------- iXBRL tag parsing ----------

' Scans raw MOPS iXBRL HTML for every <ix:nonFraction name="..." contextRef="..."
' ...>value</ix:nonFraction> tag and returns a Dictionary keyed by
' "conceptName|contextRef" -> numeric value (comma-stripped, scale-multiplied,
' sign-applied). Skips dimensional/member-breakdown contexts (contextRef
' containing "_", e.g. per-equity-component or per-segment detail) -- only
' non-dimensional consolidated totals are wanted here, same rationale as
' modSEC's segment-note handling.
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
' and order-independent (attribute order in the source varies).
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

' CDate("2026-06-01") relies on Windows regional settings to decide date-part
' order and is only safe on locales that happen to parse ISO-ish "yyyy-mm-dd"
' correctly -- DateSerial with explicit Y/M/D integers is locale-independent,
' so every "yyyy-mm" string -> Date conversion in this module goes through
' this helper rather than CDate(ym & "-01").
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
' candidates is a Variant array of concept-name strings tried in order (mirrors
' modSEC.BuildConceptMap's tagCandidates pattern for filer-to-filer tag drift).
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

' ---------- Concept extraction (ifrs-full / tifrs-* tag names -> the same metric buckets modSEC.bas uses) ----------

' Reads every metric this pipeline tracks for ONE period (durStart..durEnd for
' flow concepts, instEnd for point-in-time balance-sheet concepts) out of an
' already-parsed facts Dictionary. Called twice per Q4 quarter (once with the
' standalone quarter's duration for the quarterly series, once with the fiscal
' year's duration for the annual series) and once for every other quarter.
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
    ' TW's cash-flow-statement items use SIGNED values (confirmed live: this tag
    ' carries sign="-", rendered on the page as "(4,540,215)") since a cash flow
    ' statement sums many lines algebraically into a net change in cash -- unlike
    ' SEC's PaymentsToAcquirePropertyPlantAndEquipment, which US-GAAP defines as
    ' an inherently positive "amount paid" magnitude. The shared FCF formula
    ' (CFO - CapEx, in modCharts.BuildSnapshotTableInto and WriteTWFilingRow) is
    ' written expecting that positive convention, so this normalizes to a
    ' magnitude here -- without it, FCF came out as CFO PLUS the capex outflow
    ' instead of minus it (confirmed empirically against 3017's FY2024 figures:
    ' CFO 9,632,478 - capex magnitude 4,540,215 should give FCF 5,092,263, but
    ' the unnormalized negative value produced 14,172,693 instead).
    Dim capExRaw As Variant
    capExRaw = GetFact(facts, Array("ifrs-full:PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities"), durCtx)
    If IsNumeric(capExRaw) And CStr(capExRaw) <> "" Then capExRaw = Abs(CDbl(capExRaw))
    m("CapEx") = capExRaw
    m("Cash") = GetFact(facts, Array("ifrs-full:CashAndCashEquivalents"), instCtx)
    ' TW cash-flow statements reconcile depreciation and amortisation as two
    ' separate line items rather than SEC's single combined D&A tag -- summed
    ' here so the rest of the pipeline still sees one D&A number.
    m("Da") = SumIfNumeric(GetFact(facts, Array("ifrs-full:AdjustmentsForDepreciationExpense"), durCtx), GetFact(facts, Array("ifrs-full:AdjustmentsForAmortisationExpense"), durCtx))
    m("Cogs") = GetFact(facts, Array("tifrs-bsci-ci:OperatingCosts"), durCtx)
    ' Accounts Payable -- only needed for the Cash Conversion Cycle's DPO leg.
    m("AccountsPayable") = GetFact(facts, Array("ifrs-full:TradeAndOtherCurrentPayablesToTradeSuppliers"), instCtx)
    m("InterestExpense") = GetFact(facts, Array("ifrs-full:FinanceCosts"), durCtx)
    ' No direct "effective tax rate" tag exists in the TW taxonomy (unlike SEC's
    ' EffectiveIncomeTaxRateContinuingOperations) -- computed from the two
    ' income-statement lines that are tagged.
    m("EffectiveTaxRate") = SafeDivVariant(GetFact(facts, Array("ifrs-full:IncomeTaxExpenseContinuingOperations"), durCtx), GetFact(facts, Array("ifrs-full:ProfitLossBeforeTax", "tifrs-SCF:ProfitLossBeforeTax"), durCtx))
    ' No shares-outstanding tag is exposed on this endpoint either (SEC gets
    ' EntityCommonStockSharesOutstanding from a dedicated cover-page dei fact;
    ' MOPS's financial-statement page carries no equivalent). Backed out
    ' algebraically instead: diluted EPS = NetIncome / dilutedWeightedShares by
    ' definition, so dividing recovers the share count MOPS itself used.
    m("Shares") = SafeDivVariant(m("NetIncome"), epsV)

    ' Dividends per share: MOPS's per-quarter statement carries only an
    ' aggregate cash-dividend AMOUNT in the equity-change statement
    ' (tifrs-es:CashDividendsOfOrdinaryShare), disclosed once a year at most and
    ' not reliably present in every quarter -- not worth a fragile per-share
    ' back-out here. Left unpopulated; see BlankOutDividendRow in
    ' FetchMOPSFilings for how the resulting Dashboard cells are kept clean.

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

Private Function SafeNum(ByVal v As Variant) As Variant
    If IsNumeric(v) And CStr(v) <> "" Then
        SafeNum = CDbl(v)
    Else
        SafeNum = ""
    End If
End Function

Private Function SafeSub(ByVal a As Variant, ByVal b As Variant) As Variant
    If IsNumeric(a) And CStr(a) <> "" And IsNumeric(b) And CStr(b) <> "" Then
        SafeSub = CDbl(a) - CDbl(b)
    Else
        SafeSub = ""
    End If
End Function

' ---------- Synthetic concept maps (same shape modSEC.BuildConceptMap produces, so BuildDashboard's LookupConceptValue/LookupFyFp work unchanged) ----------

Private Function NewMapSet() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In MetricKeys()
        d.Add CStr(k), CreateObject("Scripting.Dictionary")
    Next k
    Set NewMapSet = d
End Function

' Feeds one period's already-extracted metrics into mapSet's per-metric
' accn->Collection maps, as small hand-built JSON objects carrying exactly the
' fields modSEC.LookupConceptValue/LookupFyFp read ("end"/"start"/"form"/"val"/
' "fy"/"fp"). Str(v) rather than CStr(v)/Format$(v,...) for the number literal
' deliberately -- CStr and Format$ both render using the system locale's
' decimal separator, which would corrupt this ad-hoc JSON on any machine set to
' a comma-decimal locale; Str() always uses "." regardless of locale.
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
                useStart = instEnd   ' point-in-time concepts: start==end, only "end" is ever matched on
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
                ' Set col = New Collection (not "Dim col As New Collection"): As-New
                ' only auto-instantiates when Nothing, and Dim inside a loop is one
                ' hoisted procedure-scoped variable -- reusing "As New" here would
                ' silently share ONE Collection across all 24 metric keys for this
                ' accn (confirmed: this shipped a real bug where every TW_Dashboard
                ' row for a period showed the identical value, because every metric's
                ' LookupConceptValue call searched the same merged collection and
                ' landed on whichever entry won the tie-break).
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

' ---------- TW_Filings sheet (same 48-column schema as modSEC's Filings sheet) ----------

Private Sub ClearTWFilingsSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
End Sub

Private Sub WriteTWHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    headers = Array("公司名稱", "代號 CO_ID", "表單類別", "會計年度(FY)", "會計期間(FP)", "財報截止日", "申報日期", "期別代碼 Period ID", "文件連結", _
        "營收 Revenue", "淨利 Net Income", "稀釋EPS", "總資產 Assets", "總負債 Liabilities", "營運現金流 CFO", _
        "修正版", "毛利率 Gross Margin", "營業利益率 Operating Margin", "研發費用 R&D", "SG&A", "負債比 Debt Ratio", "每股股利 Dividend/Share", _
        "申報索引頁", "部門附註連結(如有)", _
        "存貨 Inventory", "應收帳款 AR", "流動資產 Current Assets", "流動負債 Current Liabilities", "流動比率 Current Ratio", _
        "長期負債 LT Debt", "股東權益 Equity", "每股淨值 Book Value/Share", "有效稅率 Tax Rate", "資本支出 CapEx", "自由現金流 FCF", _
        "淨利率 Net Margin", "股東權益報酬率 ROE", "資產報酬率 ROA", "存貨週轉率 Inventory Turnover", "應收帳款週轉率 AR Turnover", _
        "總資產週轉率 Asset Turnover", "現金週轉率 Cash Turnover", "速動比率 Quick Ratio", "負債權益比 Debt/Equity", "權益乘數 Equity Multiplier", _
        "利息費用 Interest Expense", "利息保障倍數 Interest Coverage", _
        "事件代碼/說明 Items")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
End Sub

' Bundles one row's write-time inputs so FetchMOPSFilings can collect
' annual/quarterly rows into two ordered Collections DURING its
' newest-to-oldest fetch walk, then hand them to WriteTWFilingRow in two
' separate passes afterward (annual block first, quarterly block second) --
' matching the row order modCharts.BuildFinancialChartsInto's annualFormMarker
' split assumes (all annual rows above all quarterly rows).
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

Private Sub WriteTWFilingRow(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal spec As Object)
    Dim m As Object
    Set m = spec("metrics")

    Dim revV As Variant, gpV As Variant, opIncV As Variant, niV As Variant
    revV = m("Revenue"): gpV = m("GrossProfit"): opIncV = m("OperatingIncome"): niV = m("NetIncome")
    Dim assetsV As Variant, liabV As Variant, cfoV As Variant
    assetsV = m("Assets"): liabV = m("Liabilities"): cfoV = m("Cfo")
    Dim curAV As Variant, curLV As Variant, invV As Variant, arV As Variant, cashV As Variant
    curAV = m("CurrentAssets"): curLV = m("CurrentLiabilities"): invV = m("Inventory"): arV = m("Ar"): cashV = m("Cash")
    Dim equityV As Variant, sharesV As Variant, capexV As Variant, cogsV As Variant, intExpV As Variant
    equityV = m("StockholdersEquity"): sharesV = m("Shares"): capexV = m("CapEx"): cogsV = m("Cogs"): intExpV = m("InterestExpense")

    Dim reportDateStr As String
    reportDateStr = Format$(CDate(spec("reportDate")), "yyyy-mm-dd")

    With ws
        .Cells(rowIdx, 1).Value = spec("entityName")
        .Cells(rowIdx, 2).Value = spec("coId")
        .Cells(rowIdx, 3).Value = spec("form")
        .Cells(rowIdx, 4).Value = spec("fyLabel")
        .Cells(rowIdx, 5).Value = spec("fpLabel")
        .Cells(rowIdx, 6).Value = reportDateStr
        .Cells(rowIdx, 7).Value = reportDateStr
        .Cells(rowIdx, 8).Value = spec("fyLabel") & spec("fpLabel")
        .Cells(rowIdx, 9).Value = spec("sourceUrl")
        .Cells(rowIdx, 10).Value = SafeNum(revV)
        .Cells(rowIdx, 11).Value = SafeNum(niV)
        .Cells(rowIdx, 12).Value = SafeNum(m("Eps"))
        .Cells(rowIdx, 13).Value = SafeNum(assetsV)
        .Cells(rowIdx, 14).Value = SafeNum(liabV)
        .Cells(rowIdx, 15).Value = SafeNum(cfoV)
        .Cells(rowIdx, 16).Value = ""
        .Cells(rowIdx, 17).Value = SafeDivVariant(gpV, revV)
        .Cells(rowIdx, 18).Value = SafeDivVariant(opIncV, revV)
        .Cells(rowIdx, 19).Value = SafeNum(m("Rd"))
        .Cells(rowIdx, 20).Value = SafeNum(m("Sga"))
        .Cells(rowIdx, 21).Value = SafeDivVariant(liabV, assetsV)
        .Cells(rowIdx, 22).Value = ""
        .Cells(rowIdx, 23).Value = spec("sourceUrl")
        .Cells(rowIdx, 24).Value = ""
        .Cells(rowIdx, 25).Value = SafeNum(invV)
        .Cells(rowIdx, 26).Value = SafeNum(arV)
        .Cells(rowIdx, 27).Value = SafeNum(curAV)
        .Cells(rowIdx, 28).Value = SafeNum(curLV)
        .Cells(rowIdx, 29).Value = SafeDivVariant(curAV, curLV)
        .Cells(rowIdx, 30).Value = SafeNum(m("LongTermDebt"))
        .Cells(rowIdx, 31).Value = SafeNum(equityV)
        .Cells(rowIdx, 32).Value = SafeDivVariant(equityV, sharesV)
        .Cells(rowIdx, 33).Value = SafeNum(m("EffectiveTaxRate"))
        .Cells(rowIdx, 34).Value = SafeNum(capexV)
        .Cells(rowIdx, 35).Value = SafeSub(cfoV, capexV)
        .Cells(rowIdx, 36).Value = SafeDivVariant(niV, revV)
        .Cells(rowIdx, 37).Value = SafeDivVariant(niV, equityV)
        .Cells(rowIdx, 38).Value = SafeDivVariant(niV, assetsV)
        .Cells(rowIdx, 39).Value = SafeDivVariant(cogsV, invV)
        .Cells(rowIdx, 40).Value = SafeDivVariant(revV, arV)
        .Cells(rowIdx, 41).Value = SafeDivVariant(revV, assetsV)
        .Cells(rowIdx, 42).Value = SafeDivVariant(revV, cashV)
        .Cells(rowIdx, 43).Value = SafeDivVariant(SafeSub(curAV, invV), curLV)
        .Cells(rowIdx, 44).Value = SafeDivVariant(liabV, equityV)
        .Cells(rowIdx, 45).Value = SafeDivVariant(assetsV, equityV)
        .Cells(rowIdx, 46).Value = SafeNum(intExpV)
        .Cells(rowIdx, 47).Value = SafeDivVariant(opIncV, intExpV)
        .Cells(rowIdx, 48).Value = ""
    End With
End Sub

' ---------- Price history ----------

' Yahoo Finance's chart endpoint (already used for US tickers by modPrices.bas)
' also carries TWSE/TPEx tickers under a market-suffixed symbol -- ".TW" for
' listed (上市), ".TWO" for OTC (上櫃). There is no cheap way to know which one
' applies without an extra lookup the user chose not to require (manual CO_ID
' entry, no auxiliary ticker/market table -- see module doc comment), so this
' just tries .TW first and falls back to .TWO when that comes back empty.
Private Function FetchTWPriceHistory(ByVal coId As String, ByVal wantQ As Long) As Object
    Dim fromDate As Date
    fromDate = DateAdd("q", -(wantQ + 2), Date)

    Dim prices As Object
    Set prices = FetchDailyPrices(coId & ".TW", fromDate, Date)
    If prices.Count = 0 Then
        Set prices = FetchDailyPrices(coId & ".TWO", fromDate, Date)
    End If
    Set FetchTWPriceHistory = prices
End Function

' ---------- TW Watchlist (peer / multi-ticker comparison) ----------

' Triggered by TW_Watchlist's own Worksheet_Change (see
' SheetTWWatchlist_Code.txt) when a single CO_ID cell in the list changes.
' Mirrors modSEC.RefreshWatchlistRow's lightweight design: only the most
' recently published quarter's figures are pulled (walking back a handful of
' quarters if the very latest one isn't published yet), no multi-quarter
' history, no charts, and only the one changed row is touched.
Public Sub RefreshTWWatchlistRow(ByVal ws As Worksheet, ByVal r As Long, ByVal coIdInput As String)
    On Error GoTo ErrHandlerTWWL
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Cells(r, 2).Value = "更新中..."
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 15)).ClearContents

    Dim coId As String
    coId = Trim$(coIdInput)

    Dim curYear As Long, curSeason As Long
    Call QuarterOf(DateAdd("d", -95, Date), curYear, curSeason)

    Dim facts As Object, usedUrl As String, entityName As String
    Dim triesLeft As Long
    triesLeft = 5
    Do While triesLeft > 0
        triesLeft = triesLeft - 1
        Set facts = FetchQuarterFacts(coId, curYear, curSeason, usedUrl, entityName)
        If facts.Count > 0 Then Exit Do
        curSeason = curSeason - 1
        If curSeason = 0 Then
            curSeason = 4
            curYear = curYear - 1
        End If
    Loop

    If facts.Count = 0 Then
        ws.Cells(r, 2).Value = "找不到 MOPS 資料"
        ws.Cells(r, 16).Value = Format$(Now, "hh:nn:ss")
        GoTo CleanExitTWWL
    End If
    If entityName = "" Then entityName = coId

    Dim qStart As Date, qEnd As Date, fyStart As Date
    Call QuarterBounds(curYear, curSeason, qStart, qEnd, fyStart)

    Dim m As Object
    Set m = ExtractTWMetrics(facts, qStart, qEnd, qEnd)

    ' Prior-year same-quarter revenue for YoY growth -- reads the SAME
    ' already-fetched facts dictionary (a MOPS quarterly report tags last
    ' year's comparative period alongside this year's in the same document),
    ' so this costs no extra HTTP round-trip.
    Dim priorQStart As Date, priorQEnd As Date, priorFyStart As Date
    Call QuarterBounds(curYear - 1, curSeason, priorQStart, priorQEnd, priorFyStart)
    Dim priorRevV As Variant
    priorRevV = GetFact(facts, Array("ifrs-full:Revenue"), CtxDuration(priorQStart, priorQEnd))

    ' SafeSub/SafeDivVariant already guard every combination of missing/zero
    ' inputs and return "" instead of throwing, so this one line covers it.
    Dim revGrowth As Variant
    revGrowth = SafeDivVariant(SafeSub(m("Revenue"), priorRevV), priorRevV)

    Dim priceV As Variant
    priceV = TWLatestPrice(coId)

    Dim mktCapV As Variant
    mktCapV = ""
    If IsNumeric(priceV) And IsNumeric(m("Shares")) Then
        If CStr(priceV) <> "" And CStr(m("Shares")) <> "" Then mktCapV = CDbl(priceV) * CDbl(m("Shares"))
    End If

    Dim debtForEV As Double
    debtForEV = 0
    If IsNumeric(m("LongTermDebt")) And CStr(m("LongTermDebt")) <> "" Then debtForEV = debtForEV + CDbl(m("LongTermDebt"))
    If IsNumeric(m("ShortTermDebt")) And CStr(m("ShortTermDebt")) <> "" Then debtForEV = debtForEV + CDbl(m("ShortTermDebt"))
    Dim evV As Variant
    evV = ""
    If IsNumeric(mktCapV) And IsNumeric(m("Cash")) Then
        If CStr(mktCapV) <> "" And CStr(m("Cash")) <> "" Then evV = CDbl(mktCapV) + debtForEV - CDbl(m("Cash"))
    End If

    Dim ebitdaV As Variant
    ebitdaV = ""
    If IsNumeric(m("OperatingIncome")) And IsNumeric(m("Da")) Then
        If CStr(m("OperatingIncome")) <> "" And CStr(m("Da")) <> "" Then ebitdaV = CDbl(m("OperatingIncome")) + CDbl(m("Da"))
    End If

    Dim fcfV As Variant
    fcfV = SafeSub(m("Cfo"), m("CapEx"))

    Dim dioV As Variant, dsoV As Variant, dpoV As Variant, cccV As Variant
    dioV = SafeDaysRatio(m("Inventory"), m("Cogs"), 91)
    dsoV = SafeDaysRatio(m("Ar"), m("Revenue"), 91)
    dpoV = SafeDaysRatio(m("AccountsPayable"), m("Cogs"), 91)
    If IsNumeric(dioV) And IsNumeric(dsoV) And IsNumeric(dpoV) Then
        cccV = CDbl(dioV) + CDbl(dsoV) - CDbl(dpoV)
    Else
        cccV = ""
    End If

    Dim netMarginV As Variant, assetTurnoverV As Variant, equityMultiplierV As Variant, fcfYieldV As Variant
    netMarginV = SafeDivVariant(m("NetIncome"), m("Revenue"))
    assetTurnoverV = SafeDivVariant(m("Revenue"), m("Assets"))
    equityMultiplierV = SafeDivVariant(m("Assets"), m("StockholdersEquity"))
    fcfYieldV = SafeDivVariant(fcfV, mktCapV)

    With ws
        .Cells(r, 1).Value = coId
        .Cells(r, 2).Value = entityName
        .Cells(r, 3).Value = SafeNum(priceV)
        .Cells(r, 4).Value = SafeNum(mktCapV)
        .Cells(r, 5).Value = revGrowth
        .Cells(r, 6).Value = SafeDivVariant(priceV, m("Eps"))
        .Cells(r, 7).Value = SafeDivVariant(mktCapV, m("Revenue"))
        .Cells(r, 8).Value = SafeDivVariant(evV, ebitdaV)
        .Cells(r, 9).Value = SafeDivVariant(m("NetIncome"), m("StockholdersEquity"))
        .Cells(r, 10).Value = SafeDivVariant(m("CurrentAssets"), m("CurrentLiabilities"))
        .Cells(r, 11).Value = netMarginV
        .Cells(r, 12).Value = assetTurnoverV
        .Cells(r, 13).Value = equityMultiplierV
        .Cells(r, 14).Value = cccV
        .Cells(r, 15).Value = fcfYieldV
        .Cells(r, 16).Value = Format$(Now, "hh:nn:ss")
    End With

    ws.Cells(r, 3).NumberFormat = "#,##0.00"
    ws.Cells(r, 4).NumberFormat = "#,##0,,""M"""
    ws.Cells(r, 5).NumberFormat = "0.00%"
    ws.Range(ws.Cells(r, 6), ws.Cells(r, 8)).NumberFormat = "0.00"
    ws.Cells(r, 9).NumberFormat = "0.00%"
    ws.Cells(r, 10).NumberFormat = "0.00"
    ws.Cells(r, 11).NumberFormat = "0.00%"
    ws.Range(ws.Cells(r, 12), ws.Cells(r, 13)).NumberFormat = "0.00"
    ws.Cells(r, 14).NumberFormat = "0.0"
    ws.Cells(r, 15).NumberFormat = "0.00%"

CleanExitTWWL:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandlerTWWL:
    ws.Cells(r, 2).Value = "錯誤：" & Err.Description
    ws.Cells(r, 16).Value = Format$(Now, "hh:nn:ss")
    Resume CleanExitTWWL
End Sub

' Clears a TW_Watchlist row's fetched columns when the user deletes the
' CO_ID in column A, so a blanked-out row doesn't leave stale data behind.
Public Sub ClearTWWatchlistRow(ByVal ws As Worksheet, ByVal r As Long)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 16)).ClearContents
End Sub

Private Function TWLatestPrice(ByVal coId As String) As Variant
    Dim p As Variant
    p = LatestPrice(coId & ".TW")
    If Not (IsNumeric(p) And CStr(p) <> "") Then p = LatestPrice(coId & ".TWO")
    TWLatestPrice = p
End Function

' SafeDivVariant scaled by a day-count, for turnover-day metrics (DIO/DSO/DPO).
Private Function SafeDaysRatio(ByVal numerator As Variant, ByVal denominator As Variant, ByVal days As Double) As Variant
    Dim ratio As Variant
    ratio = SafeDivVariant(numerator, denominator)
    If IsNumeric(ratio) Then
        SafeDaysRatio = CDbl(ratio) * days
    Else
        SafeDaysRatio = ""
    End If
End Function

' ---------- Monthly revenue (台股專屬的月頻揭露，美股無此資料) ----------

' Fetches the trailing `months` months of company-level monthly revenue from
' MOPS's plain (non-XBRL) monthly-revenue table and writes a RawData block +
' dual-axis (revenue left / YoY% right) chart onto the TW dashboard sheet. Per
' CLAUDE.md's "台股標的的資料抓取" section, the wide table's "去年同月" /
' "去年同月增減%" columns have been observed to misread one column over in a
' raw scrape -- so rather than trusting the "當月營收" cell directly, this
' recomputes each month's standalone figure from consecutive cumulative-total
' cells (thisMonth = cumThisMonth - cumPriorMonth) whenever both are available,
' which is immune to that specific misalignment since it only ever reads the
' cumulative column.
Private Sub BuildMonthlyRevenue(ByVal wsIn As Worksheet, ByVal coId As String, ByVal entityName As String, ByVal months As Long)
    Dim results As Object
    Set results = CreateObject("Scripting.Dictionary")   ' "yyyy-mm" -> Array(thisRevRaw, cumRev)

    Dim cursor As Date
    cursor = DateAdd("m", -1, Date)   ' this calendar month's revenue isn't published until next month

    Dim fetched As Long
    fetched = 0
    Dim triesLeft As Long
    triesLeft = months + 3

    Do While fetched < months And triesLeft > 0
        triesLeft = triesLeft - 1
        SetTWStatus wsIn, "抓取 " & Format$(cursor, "yyyy-mm") & " 月營收中..."

        Dim rocYear As Long, mo As Long
        rocYear = Year(cursor) - 1911
        mo = Month(cursor)

        Dim thisRev As Variant, cumRev As Variant
        Dim ok As Boolean
        ok = FetchOneMonthRevenue(coId, rocYear, mo, "sii", thisRev, cumRev)
        If Not ok Then ok = FetchOneMonthRevenue(coId, rocYear, mo, "otc", thisRev, cumRev)

        If ok Then
            results(Format$(cursor, "yyyy-mm")) = Array(thisRev, cumRev)
            fetched = fetched + 1
        End If

        cursor = DateAdd("m", -1, cursor)
    Loop

    If results.Count = 0 Then Exit Sub

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

    ' NOT "Do While b >= 1 And keys(b) > tmpK" -- VBA's And does not
    ' short-circuit, so that form still evaluates keys(b) when b has already
    ' reached 0, indexing the 1-based keys() array out of bounds (confirmed:
    ' this crashed a live 36-month run with "Subscript out of range" once the
    ' insertion sort needed to shift a key all the way to position 1).
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
            ' Prior-year comparator uses the SAME recompute-from-cumulative
            ' logic, not the raw stored value, for consistency.
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

    Dim wsRaw As Worksheet
    Set wsRaw = modCharts.GetOrCreateSheet(ThisWorkbook, "TW_RawData")

    Dim rawCol As Long
    rawCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column + 1
    If wsRaw.Cells(1, 1).Value = "" Then rawCol = 1

    Dim revKeyRng As Range, revValRng As Range, yoyKeyRng As Range, yoyValRng As Range
    Call modCharts.WriteRawBlock(wsRaw, rawCol, entityName & " 月營收（近" & n & "個月）", "月份", "營收（仟元）", labelArr, revArr, n, "@", "#,##0", revKeyRng, revValRng)
    Call modCharts.WriteRawBlock(wsRaw, rawCol, entityName & " 月營收 YoY%", "月份", "YoY%", labelArr, yoyArr, n, "@", "0.00%", yoyKeyRng, yoyValRng)

    Dim wsDash As Worksheet
    Set wsDash = modCharts.GetOrCreateSheet(ThisWorkbook, "TW_Dashboard")

    Dim chartTop As Double
    chartTop = wsDash.UsedRange.Top + wsDash.UsedRange.Height + 20

    Dim bannerRow As Long
    bannerRow = CLng(chartTop / 15)
    wsDash.Cells(bannerRow, 1).Value = "═══ 月營收 Monthly Revenue（原始資料：TW_RawData 工作表）═══"
    With wsDash.Range(wsDash.Cells(bannerRow, 1), wsDash.Cells(bannerRow, 12))
        .Font.Bold = True
        .Font.Color = RGB(255, 165, 0)
        .Interior.Color = RGB(45, 45, 45)
    End With

    Dim co As ChartObject
    Set co = wsDash.ChartObjects.Add(Left:=10, Top:=chartTop + 20, Width:=750, Height:=350)
    Dim ch As Chart
    Set ch = co.Chart
    Call modCharts.ClearAutoSeries(ch)
    ch.ChartType = xlColumnClustered

    Dim serRev As Series, serYoy As Series
    Set serRev = ch.SeriesCollection.NewSeries
    serRev.Values = revValRng
    serRev.XValues = revKeyRng
    serRev.Name = "月營收"
    serRev.ChartType = xlColumnClustered

    Set serYoy = ch.SeriesCollection.NewSeries
    serYoy.Values = yoyValRng
    serYoy.XValues = yoyKeyRng
    serYoy.Name = "YoY%"
    serYoy.ChartType = xlLine
    serYoy.AxisGroup = xlSecondary

    ch.HasTitle = True
    ch.ChartTitle.Text = entityName & " — 月營收 Monthly Revenue"
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionTop

    Call modTheme.ApplyDarkThemeToChart(ch, Array(RGB(255, 140, 0), RGB(255, 180, 60)))

    Dim revLo As Double, revHi As Double, yoyLo As Double, yoyHi As Double
    If modCharts.ArrayMinMax(revArr, n, revLo, revHi) Then Call modTheme.FitValueAxis(ch.Axes(xlValue, xlPrimary), revLo, revHi)
    If modCharts.ArrayMinMax(yoyArr, n, yoyLo, yoyHi) Then Call modTheme.FitValueAxis(ch.Axes(xlValue, xlSecondary), yoyLo, yoyHi)

    Dim paintRows As Long
    paintRows = CLng((chartTop + 400) / 15) + 200
    If paintRows < 500 Then paintRows = 500
    wsDash.Range(wsDash.Cells(1, 1), wsDash.Cells(paintRows, 60)).Interior.Color = RGB(0, 0, 0)

    Call modCharts.FinishRawDataSheet(wsRaw, rawCol)
End Sub

' Parses one month's MOPS company-level revenue table
' (t21sc03_<ROCyear>_<month>.html, market = "sii" listed / "otc" OTC) for the
' row matching coId. Returns False if that company's row isn't found in this
' market/month (wrong market, or the month isn't published yet). Only reads
' the two positionally-safe columns this module actually needs -- 當月營收
' (index 1) and 累計營收 (index 6) -- rather than all ~10, since the trailing
' 備註 column is only present some months and a fixed total-column-count
' assumption would break on it.
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

' ---------- HTTP (Big5-aware) ----------

' MOPS serves Big5-encoded HTML with no charset in the HTTP Content-Type header
' (confirmed via direct inspection: the header is bare "text/html" -- charset is
' only declared in an in-document <meta> tag). WinHttpRequest's .ResponseText
' decodes using the HTTP header's charset (falling back to a Western codepage
' when none is given), which mangles every double-byte Chinese character read
' that way. Going through .ResponseBody (raw bytes) and an ADODB.Stream with an
' explicit Big5 charset override is the standard, correct fix.
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

' ---------- Sheet / status helpers ----------

Private Sub SetTWStatus(ByVal wsIn As Worksheet, ByVal msg As String)
    wsIn.Range("B1").Value = msg
    Application.StatusBar = msg
End Sub

Private Function GetTWSetting(ByVal wsIn As Worksheet, ByVal cellAddr As String, ByVal defaultVal As Long) As Long
    Dim v As Variant
    v = wsIn.Range(cellAddr).Value
    If IsNumeric(v) Then
        If CLng(v) > 0 Then
            GetTWSetting = CLng(v)
            Exit Function
        End If
    End If
    GetTWSetting = defaultVal
End Function
