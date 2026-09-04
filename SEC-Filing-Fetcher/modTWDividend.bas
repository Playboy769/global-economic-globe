Attribute VB_Name = "modTWDividend"
Option Explicit

' ---------------------------------------------------------------------------
' Taiwan dividend history, for the retention ratio / sustainable growth columns
' (modValuation) that MOPS's quarterly financial statements cannot supply.
'
' WHY A SEPARATE FETCH AT ALL: the t164sb01 financial-statement pages modMOPS
' already reads carry no per-share dividend anywhere. The equity-change
' statement has only an aggregate 現金股利 amount (tifrs-es:CashDividendsOf-
' OrdinaryShare), disclosed at most once a year and absent from most quarters,
' so retention had to stay blank for every Taiwanese company. This module fills
' that gap from MOPS's own dividend-distribution query.
'
' SOURCE: https://mopsov.twse.com.tw/mops/web/ajax_t05st09_2 (POST).
' Parameters were read off the live form rather than guessed -- note firstin
' must be the literal string "ture" (MOPS's own typo; "true"/"1" both return
' the empty page shell), and encodeURIComponent=1 must be present. TYPEK=all
' covers listed and OTC alike, so no 上市/上櫃 probe is needed: verified live
' against 3017 (上市公司) and 6488 (上櫃公司), both resolved from one request.
'
' VERIFIED AGAINST THE OFFICIAL FIGURES (2026-09-04):
'   3017 FY2025 (民國114): 盈餘 18.00 + 資本公積 3.00 = 21.00 元/股, which
'   matches openapi.twse.com.tw/v1/opendata/t187ap45_L exactly.
'
' TWO THINGS THIS SOURCE GETS RIGHT THAT THE ALTERNATIVES DON'T:
'   * It labels 股利年度 (the fiscal year whose earnings are being distributed).
'     Ex-dividend feeds (Yahoo, TWSE TWT49U) only carry the ex-date, which lags
'     the earnings year by about one year and has to be mapped by convention.
'   * It reports the DECLARED amount split by source (盈餘 / 法定盈餘公積 /
'     資本公積). Yahoo's dividend events were measured to drop the capital-
'     surplus portion entirely -- 3017's 2026 event came back as 17.898518
'     against TWSE's official 20.881604.
'
' SEMI-ANNUAL PAYERS: some filers (6488 verified) declare 上半年 and 下半年
' separately, so a year can have multiple rows. They are summed per 股利年度.
' ---------------------------------------------------------------------------

Private Const MOPS_DIV_URL As String = "https://mopsov.twse.com.tw/mops/web/ajax_t05st09_2"
Private Const TWSE_EXRIGHT_URL As String = "https://www.twse.com.tw/rwd/zh/exRight/TWT49U"

Public Const DIV_SRC_MOPS As String = "MOPS t05st09_2（申報值）"
Public Const DIV_SRC_TWSE As String = "TWSE TWT49U 除權息（權值+息值）"

' Fetches dividend history for one CO_ID over a ROC-year range.
'
' Returns a dictionary:
'   "ok"       Boolean  -- False means neither source produced anything, and
'                          callers must leave the retention ratio BLANK rather
'                          than read a missing year as "paid nothing".
'   "market"   String   -- 上市公司 / 上櫃公司, as MOPS reports it
'   "source"   String   -- which of the two sources the numbers came from
'   "cash"     Dict     -- ROC year (Long) -> cash dividend per share (Double)
'   "stock"    Dict     -- ROC year (Long) -> stock dividend per share (Double)
'   "detail"   Collection of row dictionaries, for the TW_Dividends sheet
Public Function FetchTWDividends(ByVal coId As String, ByVal fromRocYear As Long, ByVal toRocYear As Long) As Object
    Dim res As Object
    Set res = NewResult()

    Dim html As String
    On Error Resume Next
    html = HttpPostUtf8(MOPS_DIV_URL, BuildDivBody(coId, fromRocYear, toRocYear))
    On Error GoTo 0

    If Len(html) > 0 Then
        res("market") = ExtractMarket(html)
        Call ParseMopsDividendRows(html, res)
    End If

    If res("detail").Count > 0 Then
        res("ok") = True
        res("source") = DIV_SRC_MOPS
        Set FetchTWDividends = res
        Exit Function
    End If

    ' Fallback: TWSE's ex-dividend result table. Listed companies only, and it
    ' carries no 股利年度 -- the ex-date year maps to the PREVIOUS earnings year
    ' by convention (verified: 3017's 113/08/15 ex-dividend of 6.957345 is the
    ' distribution of its 民國112 earnings, declared 113/03/13 as 5.00 + 2.00).
    ' That convention breaks for the years shareholders' meetings were pushed
    ' back by regulator order, so rows from here are labelled as estimated.
    Call FetchTwseExRightFallback(coId, fromRocYear, toRocYear, res)
    If res("detail").Count > 0 Then
        res("ok") = True
        res("source") = DIV_SRC_TWSE
    End If

    Set FetchTWDividends = res
End Function

Private Function NewResult() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("market") = ""
    d("source") = ""
    Set d("cash") = CreateObject("Scripting.Dictionary")
    Set d("stock") = CreateObject("Scripting.Dictionary")
    Set d("detail") = New Collection
    Set NewResult = d
End Function

Private Function BuildDivBody(ByVal coId As String, ByVal fromRocYear As Long, ByVal toRocYear As Long) As String
    BuildDivBody = "encodeURIComponent=1&step=1&firstin=ture&off=1&keyword4=&code1=&TYPEK2=" & _
        "&checkbtn=&queryName=co_id&inpuType=co_id&TYPEK=all&isnew=" & _
        "&co_id=" & coId & "&date1=" & fromRocYear & "&date2=" & toRocYear & "&qryType=1"
End Function

' Cash dividend per share declared out of a given ROC year's earnings, or ""
' when that year has no row. Callers decide what a missing year means -- see
' the "ok" flag above.
Public Function CashDividendForRocYear(ByVal divData As Object, ByVal rocYear As Long) As Variant
    CashDividendForRocYear = ""
    If divData Is Nothing Then Exit Function
    If Not CBool(divData("ok")) Then Exit Function
    If divData("cash").Exists(rocYear) Then CashDividendForRocYear = divData("cash")(rocYear)
End Function

' True when this ROC year is old enough for a declaration to EXIST -- i.e. it
' is no later than the newest year the source actually reported.
'
' Without this, a not-yet-declared year is indistinguishable from a year the
' company chose to pay nothing, and the retention ratio silently reports 100%
' retention for every quarter of the current fiscal year (measured on 3017:
' its 2026 quarters came out at retention 1.0 and an annualised sustainable
' growth of 69%, purely because 民國115 will not be declared until 2027).
Public Function DividendYearIsDeclarable(ByVal divData As Object, ByVal rocYear As Long) As Boolean
    DividendYearIsDeclarable = False
    If divData Is Nothing Then Exit Function
    If Not CBool(divData("ok")) Then Exit Function

    Dim newest As Long
    newest = -32000
    Dim k As Variant
    For Each k In divData("cash").Keys
        If CLng(k) > newest Then newest = CLng(k)
    Next k

    DividendYearIsDeclarable = (rocYear <= newest)
End Function

Public Function StockDividendForRocYear(ByVal divData As Object, ByVal rocYear As Long) As Variant
    StockDividendForRocYear = ""
    If divData Is Nothing Then Exit Function
    If Not CBool(divData("ok")) Then Exit Function
    If divData("stock").Exists(rocYear) Then StockDividendForRocYear = divData("stock")(rocYear)
End Function

' ---------- MOPS parsing ----------

' The response uses UPPERCASE tags (<TR>/<TD>) and pads every cell with the
' literal entity &nbsp; rather than the character, both of which silently
' produce zero rows / zero amounts if not handled.
'
' Column layout (0-based), after 股利所屬年度 and its 年度/上半年/下半年 marker
' arrive merged into ONE cell:
'   0 決議進度  1 年度+期間別  2 所屬期間  3 期別  4 董事會決議日  5 股東會日期
'   6 期初未分配  7 本期淨利  8 可分配盈餘  9 分配後期末未分配
'  10 盈餘現金  11 法定公積現金  12 資本公積現金  13 現金總金額
'  14 盈餘配股  15 法定公積配股  16 資本公積配股
Private Sub ParseMopsDividendRows(ByRef html As String, ByVal res As Object)
    Dim normalized As String
    normalized = RegexReplace(html, "<tr", "<TR", True)

    Dim reCell As Object
    Set reCell = CreateObject("VBScript.RegExp")
    reCell.Global = True
    reCell.IgnoreCase = True
    reCell.Pattern = "<TD[^>]*>([\s\S]*?)</TD>"

    Dim reYear As Object
    Set reYear = CreateObject("VBScript.RegExp")
    reYear.Global = False
    reYear.Pattern = "^(\d+)年"

    Dim parts As Variant
    parts = Split(normalized, "<TR")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim mc As Object
        Set mc = reCell.Execute(CStr(parts(i)))
        If mc.Count >= 17 Then
            Dim c() As String
            ReDim c(0 To mc.Count - 1)
            Dim j As Long
            For j = 0 To mc.Count - 1
                c(j) = CleanCell(CStr(mc(j).SubMatches(0)))
            Next j

            Dim ym As Object
            Set ym = reYear.Execute(c(1))
            If ym.Count > 0 Then
                Dim rocYear As Long
                rocYear = CLng(ym(0).SubMatches(0))

                Dim cash As Double, stock As Double
                cash = Num(c(10)) + Num(c(11)) + Num(c(12))
                stock = Num(c(14)) + Num(c(15)) + Num(c(16))

                ' Semi-annual payers file one row per half-year; accumulate.
                If res("cash").Exists(rocYear) Then
                    res("cash")(rocYear) = CDbl(res("cash")(rocYear)) + cash
                Else
                    res("cash")(rocYear) = cash
                End If
                If res("stock").Exists(rocYear) Then
                    res("stock")(rocYear) = CDbl(res("stock")(rocYear)) + stock
                Else
                    res("stock")(rocYear) = stock
                End If

                Dim row As Object
                Set row = CreateObject("Scripting.Dictionary")
                row("rocYear") = rocYear
                row("periodLabel") = c(1)
                row("span") = c(2)
                row("seq") = c(3)
                row("progress") = c(0)
                row("boardDate") = c(4)
                row("agmDate") = c(5)
                row("cashEarnings") = Num(c(10))
                row("cashLegalReserve") = Num(c(11))
                row("cashCapitalSurplus") = Num(c(12))
                row("cashTotal") = cash
                row("stockEarnings") = Num(c(14))
                row("stockLegalReserve") = Num(c(15))
                row("stockCapitalSurplus") = Num(c(16))
                row("stockTotal") = stock
                row("source") = DIV_SRC_MOPS
                res("detail").Add row
            End If
        End If
    Next i
End Sub

Private Function ExtractMarket(ByRef html As String) As String
    ExtractMarket = ""
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = "本資料由[\s\S]{0,40}?(上市公司|上櫃公司|興櫃公司)"
    Dim mc As Object
    Set mc = re.Execute(html)
    If mc.Count > 0 Then ExtractMarket = CStr(mc(0).SubMatches(0))
End Function

' ---------- TWSE ex-dividend fallback ----------

Private Sub FetchTwseExRightFallback(ByVal coId As String, ByVal fromRocYear As Long, ByVal toRocYear As Long, ByVal res As Object)
    Dim y As Long
    For y = fromRocYear To toRocYear
        Dim ad As Long
        ad = y + 1911

        Dim json As String
        json = ""
        On Error Resume Next
        json = HttpGet(TWSE_EXRIGHT_URL & "?startDate=" & ad & "0101&endDate=" & ad & "1231&response=json")
        On Error GoTo 0
        If Len(json) > 0 Then Call ScanTwseExRight(json, coId, y, res)
    Next y
End Sub

' Pulls this CO_ID's rows out of one year of the ex-dividend table. Each row is
' a flat JSON array whose 2nd element is the stock code and 6th is 權值+息值;
' the ex-date year is mapped back one year to the earnings year it distributes.
Private Sub ScanTwseExRight(ByRef json As String, ByVal coId As String, ByVal exRocYear As Long, ByVal res As Object)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = """([^""]*)"",""" & coId & """,""([^""]*)"",""([^""]*)"",""([^""]*)"",""([^""]*)"",""([^""]*)"""

    Dim mc As Object
    Set mc = re.Execute(json)
    If mc.Count = 0 Then Exit Sub

    Dim earningsYear As Long
    earningsYear = exRocYear - 1

    Dim k As Long
    For k = 0 To mc.Count - 1
        Dim val As Double, kind As String
        val = Num(CStr(mc(k).SubMatches(4)))
        kind = CStr(mc(k).SubMatches(5))
        If val > 0 Then
            ' 權/息: 息 is cash only; anything else mixes in a stock component
            ' this table doesn't separate, so it is recorded but not counted as
            ' cash (the retention ratio must not treat a stock split as a payout).
            If InStr(1, kind, "息") > 0 And InStr(1, kind, "權") = 0 Then
                If res("cash").Exists(earningsYear) Then
                    res("cash")(earningsYear) = CDbl(res("cash")(earningsYear)) + val
                Else
                    res("cash")(earningsYear) = val
                End If
            End If

            Dim row As Object
            Set row = CreateObject("Scripting.Dictionary")
            row("rocYear") = earningsYear
            row("periodLabel") = earningsYear & "年 （由除息日 " & CStr(mc(k).SubMatches(0)) & " 推算）"
            row("span") = ""
            row("seq") = ""
            row("progress") = "除權息實績"
            row("boardDate") = ""
            row("agmDate") = ""
            row("cashEarnings") = IIf(InStr(1, kind, "權") = 0, val, 0)
            row("cashLegalReserve") = 0
            row("cashCapitalSurplus") = 0
            row("cashTotal") = IIf(InStr(1, kind, "權") = 0, val, 0)
            row("stockEarnings") = 0
            row("stockLegalReserve") = 0
            row("stockCapitalSurplus") = 0
            row("stockTotal") = 0
            row("source") = DIV_SRC_TWSE
            res("detail").Add row
        End If
    Next k
End Sub

' ---------- helpers ----------

Private Function HttpPostUtf8(ByVal url As String, ByVal body As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 15000, 60000
    http.Open "POST", url, False
    http.SetRequestHeader "User-Agent", SEC_USER_AGENT
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send body
    If http.Status <> 200 Then Exit Function
    ' ResponseText mis-decodes this page (the server sends no usable charset on
    ' the AJAX response), so the bytes are decoded as UTF-8 explicitly.
    HttpPostUtf8 = BytesToUtf8(http.ResponseBody)
End Function

Private Function BytesToUtf8(ByVal bytes As Variant) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                 ' adTypeBinary
    stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.Type = 2                 ' adTypeText
    stm.Charset = "utf-8"
    BytesToUtf8 = stm.ReadText
    stm.Close
End Function

Private Function CleanCell(ByVal raw As String) As String
    Dim s As String
    s = RegexReplace(raw, "<[^>]+>", " ", True)
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, ChrW(160), " ")
    s = RegexReplace(s, "\s+", " ", True)
    CleanCell = Trim$(s)
End Function

Private Function Num(ByVal s As String) As Double
    Dim t As String
    t = Trim$(Replace(s, ",", ""))
    If t = "" Then Exit Function
    If Not IsNumeric(t) Then Exit Function
    Num = CDbl(t)
End Function

Private Function RegexReplace(ByVal src As String, ByVal pattern As String, ByVal repl As String, ByVal ignoreCase As Boolean) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = ignoreCase
    re.pattern = pattern
    RegexReplace = re.Replace(src, repl)
End Function

' ---------- TW_Dividends sheet ----------

Public Sub WriteTWDividendsSheet(ByVal wb As Workbook, ByVal coId As String, ByVal entityName As String, ByVal divData As Object)
    Dim ws As Worksheet
    Set ws = modCharts.GetOrCreateSheet(wb, "TW_Dividends")
    ws.Cells.Clear

    ws.Range("A1").Value = entityName & "（" & coId & "）股利分派歷年明細　" & divData("market")
    ws.Range("A2").Value = "資料來源：" & divData("source")

    Dim headers As Variant
    headers = Array("股利年度(民國)", "所屬期間別", "所屬期間", "期別", "決議進度", "董事會決議日", "股東會日期", _
        "現金-盈餘分配(元/股)", "現金-法定盈餘公積(元/股)", "現金-資本公積(元/股)", "現金合計(元/股)", _
        "股票-盈餘轉增資(元/股)", "股票-法定公積轉增資(元/股)", "股票-資本公積轉增資(元/股)", "股票合計(元/股)", "來源")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(4, i + 1).Value = headers(i)
    Next i

    Dim r As Long
    r = 5
    Dim row As Variant
    For Each row In divData("detail")
        ws.Cells(r, 1).Value = row("rocYear")
        ws.Cells(r, 2).Value = row("periodLabel")
        ws.Cells(r, 3).Value = row("span")
        ws.Cells(r, 4).Value = row("seq")
        ws.Cells(r, 5).Value = row("progress")
        ws.Cells(r, 6).Value = row("boardDate")
        ws.Cells(r, 7).Value = row("agmDate")
        ws.Cells(r, 8).Value = row("cashEarnings")
        ws.Cells(r, 9).Value = row("cashLegalReserve")
        ws.Cells(r, 10).Value = row("cashCapitalSurplus")
        ws.Cells(r, 11).Value = row("cashTotal")
        ws.Cells(r, 12).Value = row("stockEarnings")
        ws.Cells(r, 13).Value = row("stockLegalReserve")
        ws.Cells(r, 14).Value = row("stockCapitalSurplus")
        ws.Cells(r, 15).Value = row("stockTotal")
        ws.Cells(r, 16).Value = row("source")
        r = r + 1
    Next row

    If r > 5 Then
        ws.Range(ws.Cells(5, 8), ws.Cells(r - 1, 15)).NumberFormat = "0.0000"
    End If

    Call modTheme.ApplyDarkTheme(ws, ws.Range(ws.Cells(4, 1), ws.Cells(IIf(r > 5, r - 1, 5), UBound(headers) + 1)), 1)
    ws.Range("A1:P3").Interior.Color = RGB(0, 0, 0)
    ws.Range("A1:P3").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").Font.Bold = True
    ws.Columns.AutoFit
End Sub
