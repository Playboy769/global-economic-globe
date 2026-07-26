Attribute VB_Name = "Attach"
Option Explicit

' ================================================================
' SHARED FUNCTIONS MODULE
' Used by: PortfolioDashboard_v3, UserForms, Correlation sheet
' ================================================================

Public Type TechIndicators
    price As Double
    ChangePercent As Double
    Change180D As Double
    DistFromHigh As Double
    Bias5 As Double
    Bias20 As Double
    Bias60 As Double
    Bias120 As Double
    Bias240 As Double
    PERatio As Double
    ForwardPE As Double
    PEGRatio As Double
    EPSNextQ As Double
    EPSNextY As Double
    Success As Boolean
End Type

' ================================================================
' PRICE CACHE (in-memory, 5-minute TTL)
' ================================================================
Public g_PriceCache As Object
Public g_CacheTime As Date

Function GetCompanyName(Ticker As String) As String
    Dim url As String, http As Object, response As String
    Dim nameStr As String, startPos As Long, endPos As Long

    Ticker = UCase(Trim(Ticker))
    If InStr(Ticker, ".TW") = 0 And InStr(Ticker, ".TWO") = 0 And IsNumeric(Ticker) Then
        Ticker = Ticker & ".TW"
    End If

    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=1d"

    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo 0

    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        response = .responseText
    End With

    Dim searchTag As String
    searchTag = """longName"":"""
    startPos = InStr(response, searchTag)

    If startPos > 0 Then
        startPos = startPos + Len(searchTag)
        endPos = InStr(startPos, response, """")
        GetCompanyName = Mid(response, startPos, endPos - startPos)
    Else
        searchTag = """shortName"":"""
        startPos = InStr(response, searchTag)
        If startPos > 0 Then
            startPos = startPos + Len(searchTag)
            endPos = InStr(startPos, response, """")
            GetCompanyName = Mid(response, startPos, endPos - startPos)
        Else
            GetCompanyName = Ticker
        End If
    End If
    Set http = Nothing
End Function

Function GetCurrencyType(Ticker As String) As String
    Ticker = UCase(Trim(Ticker))
    If InStr(Ticker, ".TW") > 0 Or InStr(Ticker, ".TWO") > 0 Then
        GetCurrencyType = "TWD"
    Else
        GetCurrencyType = "USD"
    End If
End Function

' ================================================================
' Technical Analysis
' ================================================================
Function GetTechData(Ticker As String) As TechIndicators
    Dim http As Object, response As String
    Dim result As TechIndicators
    Dim closeArr() As Double
    Dim i As Long, count As Long
    Dim currentPrice As Double, maxHigh As Double
    Dim rawTicker As String

    rawTicker = UCase(Trim(Ticker))
    result.Success = False
    Ticker = rawTicker

    If InStr(Ticker, ".TW") = 0 And InStr(Ticker, ".TWO") = 0 And IsNumeric(Ticker) Then
        Ticker = Ticker & ".TW"
    End If

    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo 0

    response = FetchYahooData(http, Ticker, "chart")

    If InStr(response, """timestamp""") = 0 And IsNumeric(rawTicker) Then
        Ticker = rawTicker & ".TWO"
        response = FetchYahooData(http, Ticker, "chart")
    End If

    Dim closeStart As Long, closeEnd As Long, closeStr As String, rawArr() As String
    Dim qPos As Long
    qPos = InStr(response, """quote"":[{")
    If qPos = 0 Then qPos = 1

    closeStart = InStr(qPos, response, """close"":[") + 9
    closeEnd = InStr(closeStart, response, "]")

    If closeStart > 8 And closeEnd > closeStart Then
        closeStr = Mid(response, closeStart, closeEnd - closeStart)
        rawArr = Split(closeStr, ",")
        ReDim closeArr(0 To UBound(rawArr))
        count = 0: maxHigh = 0

        For i = 0 To UBound(rawArr)
            If IsNumeric(rawArr(i)) Then
                closeArr(count) = Val(rawArr(i))
                If closeArr(count) > maxHigh Then maxHigh = closeArr(count)
                count = count + 1
            End If
        Next i

        If count > 0 Then
            ReDim Preserve closeArr(0 To count - 1)
            currentPrice = closeArr(count - 1)

            Dim metaPos As Long: metaPos = InStr(response, """regularMarketPrice"":")
            If metaPos > 0 Then
                Dim vals As Long: vals = metaPos + 21
                Dim valE As Long: valE = InStr(vals, response, ",")
                If valE = 0 Then valE = InStr(vals, response, "}")
                If valE > vals Then
                    Dim metaVal As String: metaVal = Trim(Mid(response, vals, valE - vals))
                    If IsNumeric(metaVal) And Val(metaVal) > 0 Then
                        Dim metaPx As Double: metaPx = Val(metaVal)
                        If metaPx <> currentPrice Then
                            currentPrice = metaPx
                            closeArr(count - 1) = currentPrice
                            If currentPrice > maxHigh Then maxHigh = currentPrice
                        End If
                    End If
                End If
            End If

            result.price = currentPrice

            If count >= 2 Then
                Dim prev As Double: prev = closeArr(count - 2)
                If prev > 0 Then result.ChangePercent = (currentPrice - prev) / prev
            End If

            Dim idx180 As Long: idx180 = count - 126
            If idx180 >= LBound(closeArr) And idx180 <= UBound(closeArr) Then
                If closeArr(idx180) > 0 Then
                    result.Change180D = (currentPrice - closeArr(idx180)) / closeArr(idx180)
                End If
            End If

            If maxHigh > 0 Then result.DistFromHigh = (currentPrice - maxHigh) / maxHigh

            result.Bias5 = CalculateBias(closeArr, 5, currentPrice)
            result.Bias20 = CalculateBias(closeArr, 20, currentPrice)
            result.Bias60 = CalculateBias(closeArr, 60, currentPrice)
            result.Bias120 = CalculateBias(closeArr, 120, currentPrice)
            result.Bias240 = CalculateBias(closeArr, 240, currentPrice)
            result.Success = True

            Call GetFundamentalData(http, Ticker, result)
        End If
    End If
    GetTechData = result
End Function

' ================================================================
' GET STOCK PRICE (with 5-min cache)
'
' Routing:
'   .TW / .TWO  -> TWSE / TPEx OpenAPI  (TaiwanPriceFetcher module)
'   US / others -> Yahoo Finance         (original path, Yahoo fallback)
'
' Adding routing here fixes ALL callers automatically:
'   PortfolioDashboard_v3, TickerInsight, CampaignSheet, etc.
' ================================================================
Public Function GetStockPrice(Ticker As String) As Double
    Ticker = UCase(Trim(Ticker))

    If g_PriceCache Is Nothing Then
        Set g_PriceCache = CreateObject("Scripting.Dictionary")
        g_CacheTime = Now
    End If

    If Now - g_CacheTime > TimeSerial(0, 5, 0) Then
        g_PriceCache.RemoveAll
        g_CacheTime = Now
    End If

    If g_PriceCache.Exists(Ticker) Then
        GetStockPrice = g_PriceCache(Ticker)
        Exit Function
    End If

    ' ── STEP 1: Normalise ticker ─────────────────────────────────────────
    Dim fullTicker As String: fullTicker = Ticker
    If InStr(fullTicker, ".TW") = 0 And InStr(fullTicker, ".TWO") = 0 And IsNumeric(fullTicker) Then
        fullTicker = fullTicker & ".TW"
    End If

    ' ── STEP 2: Yahoo Finance (primary) ──────────────────────────────────
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo 0

    Dim safeTicker As String: safeTicker = Replace(fullTicker, "^", "%5E")
    Dim url As String
    Dim resp As String: resp = ""
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & safeTicker & "?interval=1d&range=5d"

    On Error Resume Next
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .setRequestHeader "Accept", "application/json"
        .send
        resp = .responseText
    End With
    On Error GoTo 0

    Dim px As Double: px = ParseClosePrice(resp)

    ' .TWO fallback for bare numeric tickers (still within Yahoo)
    If px = 0 And IsNumeric(Ticker) And InStr(fullTicker, ".TWO") = 0 Then
        url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & ".TWO?interval=1d&range=5d"
        On Error Resume Next
        With http
            .Open "GET", url, False
            .setRequestHeader "User-Agent", "Mozilla/5.0"
            .send
            resp = .responseText
        End With
        On Error GoTo 0
        px = ParseClosePrice(resp)
        If px > 0 Then fullTicker = Ticker & ".TWO"
    End If

    ' ── STEP 3: TWSE / TPEx OpenAPI (backup when Yahoo returns 0) ────────
    If px = 0 And InStr(fullTicker, ".TW") > 0 Then
        Dim openApiPx As Double: openApiPx = GetTWStockPrice(fullTicker)
        If openApiPx > 0 Then
            g_PriceCache(Ticker) = openApiPx
            GetStockPrice = openApiPx
            Set http = Nothing
            Exit Function
        End If
    End If
    ' ─────────────────────────────────────────────────────────────────────

    If px > 0 Then g_PriceCache(Ticker) = px
    GetStockPrice = px
    Set http = Nothing
End Function

' ----------------------------------------------------------------
Private Function ParseClosePrice(resp As String) As Double
    ParseClosePrice = 0

    Dim metaPos As Long: metaPos = InStr(resp, """regularMarketPrice"":")
    If metaPos > 0 Then
        Dim vals As Long: vals = metaPos + 21
        Dim valE As Long: valE = InStr(vals, resp, ",")
        If valE = 0 Then valE = InStr(vals, resp, "}")
        If valE > vals Then
            Dim metaVal As String: metaVal = Trim(Mid(resp, vals, valE - vals))
            If IsNumeric(metaVal) And Val(metaVal) > 0 Then
                ParseClosePrice = Val(metaVal)
                Exit Function
            End If
        End If
    End If

    Dim csPos As Long: csPos = InStr(resp, """close"":[")
    If csPos = 0 Then Exit Function
    Dim csS As Long: csS = csPos + 9
    Dim csE As Long: csE = InStr(csS, resp, "]")
    If csE <= csS Then Exit Function

    Dim parts() As String: parts = Split(Mid(resp, csS, csE - csS), ",")
    Dim i As Long
    For i = UBound(parts) To 0 Step -1
        Dim pt As String: pt = Trim(parts(i))
        If IsNumeric(pt) Then
            If CDbl(pt) > 0 Then
                ParseClosePrice = CDbl(pt)
                Exit Function
            End If
        End If
    Next i
End Function

Sub GetFundamentalData(http As Object, Ticker As String, ByRef res As TechIndicators)
    Dim response As String
    response = FetchYahooData(http, Ticker, "quoteSummary")
    If Len(response) < 100 Then Exit Sub

    response = Replace(Replace(Replace(Replace(response, " ", ""), vbCr, ""), vbLf, ""), vbTab, "")

    res.PERatio = JsonExtract(response, "trailingPE")
    res.ForwardPE = JsonExtract(response, "forwardPE")
    res.PEGRatio = JsonExtract(response, "pegRatio")
    res.EPSNextQ = JsonExtractTrend(response, 1)
    res.EPSNextY = JsonExtractTrend(response, 3)

    Dim apiPrice As Double: apiPrice = JsonExtract(response, "regularMarketPrice")
    Dim apiPrev As Double: apiPrev = JsonExtract(response, "regularMarketPreviousClose")
    If apiPrice > 0 And apiPrev > 0 Then
        res.ChangePercent = (apiPrice - apiPrev) / apiPrev
    End If
End Sub

Public Function FetchYahooData(http As Object, Ticker As String, mode As String) As String
    Dim url As String
    If mode = "chart" Then
        url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=2y"
    Else
        url = "https://query1.finance.yahoo.com/v10/finance/quoteSummary/" & Ticker & _
              "?modules=summaryDetail,defaultKeyStatistics,earningsTrend,price"
    End If
    On Error Resume Next
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        FetchYahooData = .responseText
    End With
    On Error GoTo 0
End Function

Function CalculateBias(priceArr() As Double, period As Long, currentPrice As Double) As Double
    If UBound(priceArr) + 1 < period Then CalculateBias = 0: Exit Function
    Dim k As Double: k = 2 / (period + 1)
    Dim ema As Double: ema = priceArr(0)
    Dim i As Long
    For i = 1 To UBound(priceArr)
        ema = (priceArr(i) - ema) * k + ema
    Next i
    If ema > 0 Then CalculateBias = (currentPrice - ema) / ema Else CalculateBias = 0
End Function

Function JsonExtract(json As String, key As String) As Double
    Dim searchStr As String: searchStr = """" & key & """:{""raw"":"
    Dim pos As Long: pos = InStr(json, searchStr)
    If pos = 0 Then JsonExtract = 0: Exit Function

    Dim valStart As Long: valStart = pos + Len(searchStr)
    Dim valEnd As Long: valEnd = InStr(valStart, json, ",")
    If valEnd = 0 Then valEnd = InStr(valStart, json, "}")
    If valEnd > valStart Then
        Dim tmp As String: tmp = Mid(json, valStart, valEnd - valStart)
        If IsNumeric(tmp) Then JsonExtract = Val(tmp)
    End If
End Function

Function JsonExtractTrend(json As String, blockIndex As Long) As Double
    Dim pos As Long: pos = 1
    Dim i As Long
    For i = 0 To blockIndex
        pos = InStr(pos, json, """earningsEstimate"":")
        If pos = 0 Then JsonExtractTrend = 0: Exit Function
        pos = pos + Len("""earningsEstimate"":")
    Next i
    Dim searchStr As String: searchStr = """avg"":{""raw"":"
    pos = InStr(pos, json, searchStr)
    If pos > 0 Then
        Dim valStart As Long: valStart = pos + Len(searchStr)
        Dim valEnd As Long: valEnd = InStr(valStart, json, ",")
        If valEnd = 0 Then valEnd = InStr(valStart, json, "}")
        If valEnd > valStart Then
            Dim tmp As String: tmp = Mid(json, valStart, valEnd - valStart)
            If IsNumeric(tmp) Then JsonExtractTrend = Val(tmp)
        End If
    End If
End Function

' ================================================================
' Beta Calculation
' ================================================================
Function GetBenchmarkTicker(Sector As String, stockTicker As String) As String
    If InStr(UCase(stockTicker), ".TW") > 0 Or InStr(UCase(stockTicker), ".TWO") > 0 Then
        GetBenchmarkTicker = "^TWII": Exit Function
    End If
    Select Case Sector
    Case ChrW(&H8CC7) & ChrW(&H8A0A) & ChrW(&H79D1) & ChrW(&H6280) & " (Information Technology)": GetBenchmarkTicker = "^IXIC"
    Case ChrW(&H5DE5) & ChrW(&H696D) & " (Industrials)": GetBenchmarkTicker = "^DJI"
    Case ChrW(&H91D1) & ChrW(&H878D) & " (Financials)": GetBenchmarkTicker = "XLF"
    Case ChrW(&H91AB) & ChrW(&H7642) & ChrW(&H4FDD) & ChrW(&H5065) & " (Health Care)": GetBenchmarkTicker = "XLV"
    Case ChrW(&H80FD) & ChrW(&H6E90) & " (Energy)": GetBenchmarkTicker = "XLE"
    Case Else: GetBenchmarkTicker = "^GSPC"
    End Select
End Function

Function CalculateBeta30D(stockTicker As String, marketTicker As String) As Double
    Dim stockPrices As Object, marketPrices As Object
    Set stockPrices = GetHistoryPrices(stockTicker)
    Set marketPrices = GetHistoryPrices(marketTicker)

    If stockPrices.Count < 5 Or marketPrices.Count < 5 Then
        CalculateBeta30D = -999: Exit Function
    End If

    Dim keys As Variant: keys = stockPrices.Keys
    Dim totalKeys As Long: totalKeys = UBound(keys) + 1
    Dim startIdx As Long: startIdx = IIf(totalKeys > 30, totalKeys - 30, 0)

    Dim arrStock() As Double
    Dim arrMarket() As Double
    ReDim arrStock(1 To totalKeys)
    ReDim arrMarket(1 To totalKeys)

    Dim validPairs As Long
    Dim prevStock As Double, prevMarket As Double
    Dim i As Long, dateKey As Variant

    For i = startIdx To UBound(keys)
        dateKey = keys(i)
        If marketPrices.Exists(dateKey) Then
            Dim currStock As Double: currStock = stockPrices(dateKey)
            Dim currMarket As Double: currMarket = marketPrices(dateKey)
            If prevStock > 0 Then
                validPairs = validPairs + 1
                arrStock(validPairs) = (currStock - prevStock) / prevStock
                arrMarket(validPairs) = (currMarket - prevMarket) / prevMarket
            End If
            prevStock = currStock: prevMarket = currMarket
        End If
    Next i

    If validPairs >= 10 Then
        Dim cleanS() As Double, cleanM() As Double
        ReDim cleanS(1 To validPairs): ReDim cleanM(1 To validPairs)
        For i = 1 To validPairs
            cleanS(i) = arrStock(i): cleanM(i) = arrMarket(i)
        Next i
        On Error Resume Next
        CalculateBeta30D = Application.WorksheetFunction.Slope(cleanS, cleanM)
        If Err.Number <> 0 Then CalculateBeta30D = -999
        On Error GoTo 0
    Else
        CalculateBeta30D = -999
    End If
End Function

Function GetHistoryPrices(Ticker As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Ticker = UCase(Trim(Ticker))
    If InStr(Ticker, ".TW") = 0 And InStr(Ticker, ".TWO") = 0 And IsNumeric(Ticker) Then
        Ticker = Ticker & ".TW"
    End If

    Dim http As Object, response As String
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?range=1y&interval=1d", False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        response = .responseText
    End With
    On Error GoTo 0

    Dim tsStart As Long: tsStart = InStr(response, """timestamp"":[") + 13
    Dim tsEnd As Long: tsEnd = InStr(tsStart, response, "]")
    Dim csStart As Long: csStart = InStr(response, """adjclose"":[") + 12
    Dim csEnd As Long: csEnd = InStr(csStart, response, "]")

    If tsStart > 12 And tsEnd > tsStart And csStart > 11 And csEnd > csStart Then
        Dim tsArr() As String: tsArr = Split(Mid(response, tsStart, tsEnd - tsStart), ",")
        Dim closeArr() As String: closeArr = Split(Mid(response, csStart, csEnd - csStart), ",")
        If UBound(tsArr) = UBound(closeArr) Then
            Dim i As Long
            For i = 0 To UBound(tsArr)
                If IsNumeric(tsArr(i)) And IsNumeric(closeArr(i)) Then
                    Dim dv As Date: dv = Int(CLng(tsArr(i)) / 86400 + 25569)
                    If Not dict.Exists(dv) Then dict.Add dv, CDbl(closeArr(i))
                End If
            Next i
        End If
    End If
    Set GetHistoryPrices = dict
End Function

Sub CalculateRealizedPnL()
    Dim wsTrans As Worksheet, wsReal As Worksheet, wsPort As Worksheet
    Set wsTrans = ThisWorkbook.Sheets("Transactions")
    Set wsReal = ThisWorkbook.Sheets("Realized")
    Set wsPort = ThisWorkbook.Sheets("RR4")
    Application.ScreenUpdating = False
    wsReal.Range("A2:I10000").ClearContents

    Dim hdrs As Variant
    hdrs = Array("TICKER", "PNL", "SHARES", "AVG COST", "NET AMT", "STRATEGY", "PNL(TWD)", "DATE", "BROKER")
    Dim hc As Integer
    For hc = 0 To 8
        wsReal.Cells(1, hc + 1).Value = hdrs(hc)
        wsReal.Cells(1, hc + 1).Font.Bold = True
    Next hc

    Dim dictLots As Object: Set dictLots = CreateObject("Scripting.Dictionary")

    Dim exRate As Double
    On Error Resume Next: exRate = Val(wsPort.Range("B2").Value): On Error GoTo 0
    If exRate <= 0 Then exRate = 32

    Dim lastRow As Long: lastRow = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
    Dim writeRow As Long: writeRow = 2
    Dim rr As Long

    For rr = 2 To lastRow
        Dim tDate As Variant: tDate = wsTrans.Cells(rr, 2).Value
        Dim Ticker As String: Ticker = UCase(Trim(wsTrans.Cells(rr, 3).Value))
        Dim Action As String: Action = UCase(Trim(wsTrans.Cells(rr, 4).Value))
        Dim shares As Double: shares = Val(wsTrans.Cells(rr, 5).Value)
        Dim NetAmount As Double: NetAmount = Val(wsTrans.Cells(rr, 9).Value)
        Dim strat As String: strat = CStr(wsTrans.Cells(rr, 12).Value)
        Dim broker As String: broker = Trim(CStr(wsTrans.Cells(rr, 14).Value))
        If broker = "" Then broker = "Default"

        Dim currType As String
        If InStr(Ticker, ".TW") > 0 Or InStr(Ticker, ".TWO") > 0 Then
            currType = "TWD"
        Else
            currType = "USD"
        End If

        If Ticker <> "" And shares > 0 Then
            Dim posKey As String: posKey = Ticker & "|" & broker

            If Not dictLots.Exists(posKey) Then
                Dim newCol As Collection
                Set newCol = New Collection
                dictLots.Add posKey, newCol
            End If

            Select Case Action
            Case "BUY"
                Dim costPerShare As Double
                costPerShare = NetAmount / shares
                dictLots(posKey).Add Array(shares, costPerShare)

            Case "ADJUSTCOST"
                Dim totalSh As Double: totalSh = 0
                Dim ll As Long
                For ll = 1 To dictLots(posKey).Count
                    totalSh = totalSh + dictLots(posKey)(ll)(0)
                Next ll
                If totalSh > 0 Then
                    Dim adjPerSh As Double: adjPerSh = NetAmount / totalSh
                    For ll = 1 To dictLots(posKey).Count
                        Dim lotArr As Variant: lotArr = dictLots(posKey)(ll)
                        lotArr(1) = lotArr(1) + adjPerSh
                        If ll < dictLots(posKey).Count Then
                            dictLots(posKey).Add lotArr, Before:=ll + 1
                            dictLots(posKey).Remove ll
                        Else
                            dictLots(posKey).Add lotArr
                            dictLots(posKey).Remove ll
                        End If
                    Next ll
                End If

            Case "SELL"
                Dim remaining As Double: remaining = shares
                Dim totalCostBasis As Double: totalCostBasis = 0

                Do While remaining > 0.00001 And dictLots(posKey).Count > 0
                    Dim frontLot As Variant: frontLot = dictLots(posKey)(1)
                    Dim lotSh As Double: lotSh = frontLot(0)
                    Dim lotCps As Double: lotCps = frontLot(1)

                    If lotSh <= remaining + 0.00001 Then
                        totalCostBasis = totalCostBasis + lotSh * lotCps
                        remaining = remaining - lotSh
                        dictLots(posKey).Remove 1
                    Else
                        totalCostBasis = totalCostBasis + remaining * lotCps
                        Dim updatedLot As Variant
                        updatedLot = Array(lotSh - remaining, lotCps)
                        If dictLots(posKey).Count > 1 Then
                            dictLots(posKey).Add updatedLot, Before:=2
                        Else
                            dictLots(posKey).Add updatedLot
                        End If
                        dictLots(posKey).Remove 1
                        remaining = 0
                    End If
                Loop

                Dim fifoAvgC As Double
                If shares > 0 Then fifoAvgC = totalCostBasis / shares Else fifoAvgC = 0
                Dim rlPnl As Double: rlPnl = NetAmount - totalCostBasis
                Dim rlTWD As Double: rlTWD = IIf(currType = "USD", rlPnl * exRate, rlPnl)

                Dim finalDate As Date
                If IsDate(tDate) Then finalDate = CDate(tDate) Else finalDate = Date

                With wsReal
                    .Cells(writeRow, 1) = Ticker
                    .Cells(writeRow, 2) = rlPnl: .Cells(writeRow, 2).NumberFormat = "#,##0.00"
                    .Cells(writeRow, 3) = shares
                    .Cells(writeRow, 4) = fifoAvgC
                    .Cells(writeRow, 5) = NetAmount
                    .Cells(writeRow, 6) = strat
                    .Cells(writeRow, 7) = rlTWD: .Cells(writeRow, 7).NumberFormat = "#,##0"
                    .Cells(writeRow, 8) = finalDate: .Cells(writeRow, 8).NumberFormat = "yyyy/m/d"
                    .Cells(writeRow, 9) = broker
                    If rlTWD > 0 Then
                        .Cells(writeRow, 2).Font.Color = RGB(255, 100, 100)
                        .Cells(writeRow, 7).Font.Color = RGB(255, 100, 100)
                    Else
                        .Cells(writeRow, 2).Font.Color = RGB(0, 255, 100)
                        .Cells(writeRow, 7).Font.Color = RGB(0, 255, 100)
                    End If
                End With
                writeRow = writeRow + 1

            End Select
        End If
    Next rr

    ' Update RR4 Entry PX
    Dim portLastRow As Long: portLastRow = wsPort.Cells(wsPort.Rows.Count, "A").End(xlUp).Row
    Dim pr As Long
    For pr = 5 To portLastRow
        Dim pTicker As String: pTicker = UCase(Trim(wsPort.Cells(pr, 1).Value))
        Dim pBroker As String: pBroker = Trim(CStr(wsPort.Cells(pr, 13).Value))
        If pBroker = "" Then pBroker = "Default"

        Dim pKey As String: pKey = pTicker & "|" & pBroker

        If dictLots.Exists(pKey) And pTicker <> "" Then
            Dim totSh As Double: totSh = 0
            Dim totCost As Double: totCost = 0
            For ll = 1 To dictLots(pKey).Count
                totSh = totSh + dictLots(pKey)(ll)(0)
                totCost = totCost + dictLots(pKey)(ll)(0) * dictLots(pKey)(ll)(1)
            Next ll
            If totSh > 0.00001 Then
                wsPort.Cells(pr, 9).Value = totCost / totSh
            End If
        End If
    Next pr

    wsReal.Columns("A:I").AutoFit
    Application.ScreenUpdating = True
End Sub

Sub ClearAllData()
    If MsgBox("Clear ALL data? This cannot be undone!", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub

    Dim wsTrans As Worksheet, wsPort As Worksheet, wsReal As Worksheet, wsHist As Worksheet
    On Error Resume Next
    Set wsTrans = ThisWorkbook.Sheets("Transactions")
    Set wsPort = ThisWorkbook.Sheets("RR4")
    Set wsReal = ThisWorkbook.Sheets("Realized")
    Set wsHist = ThisWorkbook.Sheets("HistoryLog")
    On Error GoTo 0

    Dim LR As Long
    If Not wsTrans Is Nothing Then
        LR = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
        If LR > 1 Then wsTrans.Range("A2:N" & LR).ClearContents
    End If
    If Not wsReal Is Nothing Then
        LR = wsReal.Cells(wsReal.Rows.Count, "A").End(xlUp).Row
        If LR > 1 Then wsReal.Range("A2:I" & LR).ClearContents
    End If
    If Not wsHist Is Nothing Then
        LR = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row
        If LR > 1 Then wsHist.Range("A2:G" & LR).ClearContents
    End If

    MsgBox "All data cleared.", vbInformation
End Sub

' ================================================================
' Form Launchers
' ================================================================
Sub OpenTransactionForm()
    frmTransaction.Show
End Sub

Sub OpenDeleteForm()
    frmDeleteTransaction.Show
End Sub

' ================================================================
' Realized Sheet Filter UI
' ================================================================
Sub InitFilterInterface()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Realized")
    On Error Resume Next: ws.Unprotect: On Error GoTo 0
    ws.AutoFilterMode = False
    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Size = 12
    End With
    With ws.Range("J1:L20")
        .UnMerge: .ClearContents: .ColumnWidth = 12
    End With
    With ws
        .Range("J1:L6").Interior.Color = RGB(30, 30, 30)
        .Range("J1:L1").Merge
        .Range("J1").HorizontalAlignment = xlCenter
        .Range("J1").Font.Bold = True
        .Range("J1").Font.Color = RGB(255, 215, 0)
        .Range("J2").Value = "Start Date:"
        .Range("J3").Value = "End Date:"
        .Range("K2").Value = DateSerial(Year(Date), 1, 1)
        .Range("K3").Value = Date
        With .Range("K2:K3")
            .NumberFormat = "yyyy/m/d"
            .Interior.Color = RGB(60, 60, 60)
            .Borders.LineStyle = xlContinuous
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
        .Range("J5").Value = "Period P&L:"
        .Range("J5").Font.Bold = True
        .Range("K5").NumberFormat = "$#,##0"
        .Range("K5").Font.Size = 14
        .Range("K5").Font.Bold = True
        .Range("L2").Value = "(e.g. 2025/1/1)"
        .Range("L2").Font.Color = RGB(150, 150, 150)
        .Range("L2").Font.Size = 10
    End With
End Sub

Sub FilterRealizedData()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Realized")
    Dim LR As Long: LR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If LR < 2 Then MsgBox "No data to filter", vbInformation: Exit Sub

    Dim sd As Date, ed As Date
    On Error Resume Next
    If IsDate(ws.Range("K2").Value) Then sd = CDate(ws.Range("K2").Value)
    If IsDate(ws.Range("K3").Value) Then ed = CDate(ws.Range("K3").Value)
    On Error GoTo 0
    If sd = 0 Or ed = 0 Then MsgBox "Enter valid dates in K2 and K3", vbExclamation: Exit Sub

    ws.AutoFilterMode = False
    ws.Range("A1:I" & LR).AutoFilter Field:=8, Criteria1:=">=" & sd, Operator:=xlAnd, Criteria2:="<=" & ed

    Dim PnL As Double
    On Error Resume Next
    PnL = Application.WorksheetFunction.Subtotal(109, ws.Range("G2:G" & LR))
    On Error GoTo 0

    ws.Range("K5").Value = PnL
    ws.Range("K5").Font.Color = IIf(PnL >= 0, RGB(255, 100, 100), RGB(0, 255, 100))
    MsgBox "Filter applied!", vbInformation
End Sub

Sub ClearRealizedFilter()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Realized")
    ws.AutoFilterMode = False
    Dim LR As Long: LR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim total As Double
    If LR >= 2 Then total = Application.WorksheetFunction.Sum(ws.Range("G2:G" & LR))
    ws.Range("K5").Value = total
    ws.Range("K5").Font.Color = IIf(total >= 0, RGB(255, 100, 100), RGB(0, 255, 100))
End Sub

' ================================================================
' GitHub Coverage Report
' ================================================================
Sub QueryCoverageReport()
    Dim Ticker As String
    Ticker = InputBox("Enter TW stock code (e.g. 2330):", "Coverage Report")
    Ticker = Trim(Ticker)
    If Ticker = "" Then Exit Sub
    Dim pureNum As String
    pureNum = Replace(Replace(UCase(Ticker), ".TWO", ""), ".TW", "")
    If Not IsNumeric(pureNum) Then MsgBox "Numeric code only (e.g. 2330)", vbExclamation: Exit Sub

    Application.StatusBar = "Searching " & pureNum & "..."
    Application.ScreenUpdating = False

    Dim filePath As String: filePath = SearchGitHubFile(pureNum)
    If filePath = "" Then
        Application.StatusBar = False: Application.ScreenUpdating = True
        MsgBox "Report not found for " & pureNum, vbExclamation: Exit Sub
    End If

    Dim content As String: content = FetchRawContent(filePath)
    If content = "" Then
        Application.StatusBar = False: Application.ScreenUpdating = True
        MsgBox "Failed to load report.", vbExclamation: Exit Sub
    End If

    Call RenderReportToSheet(content, pureNum)
    Application.StatusBar = pureNum & " loaded."
    Application.ScreenUpdating = True
End Sub

Function EncodeURL(text As String) As String
    Dim i As Long, result As String, c As String
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        Select Case c
        Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", ".", "/", "~"
            result = result & c
        Case Else
            Dim bytes() As Byte: bytes = StrConv(c, vbFromUnicode)
            Dim b As Long
            For b = 0 To UBound(bytes)
                result = result & "%" & UCase(Hex(bytes(b)))
            Next b
        End Select
    Next i
    EncodeURL = result
End Function

Function SearchGitHubFile(tickerNum As String) As String
    Dim http As Object, response As String
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim ghToken As String: ghToken = GetGitHubToken()
    On Error Resume Next
    With http
        .Open "GET", "https://api.github.com/repos/Timeverse/My-TW-Coverage/git/trees/master?recursive=1", False
        .setRequestHeader "User-Agent", "VBA-Excel-Client"
        .setRequestHeader "Accept", "application/vnd.github.v3+json"
        If ghToken <> "" Then .setRequestHeader "Authorization", "token " & ghToken
        .send
        response = .responseText
    End With
    On Error GoTo 0
    If Len(response) < 100 Then SearchGitHubFile = "": Exit Function

    Dim pat As String: pat = """path"":""Pilot_Reports"
    Dim pos As Long: pos = 1
    Do
        pos = InStr(pos, response, pat)
        If pos = 0 Then SearchGitHubFile = "": Exit Function
        Dim vs As Long: vs = pos + 8
        Dim ve As Long: ve = InStr(vs, response, """")
        If ve > vs Then
            Dim pv As String: pv = Mid(response, vs, ve - vs)
            If InStr(pv, "/" & tickerNum & "_") > 0 Then
                SearchGitHubFile = pv: Exit Function
            End If
        End If
        pos = pos + Len(pat)
    Loop
End Function

Function GetGitHubToken() As String
    On Error Resume Next
    GetGitHubToken = ThisWorkbook.Sheets("HistoryLog").Range("Z3").Value
    On Error GoTo 0
End Function

Sub SetGitHubToken()
    Dim token As String
    token = InputBox("Paste GitHub Token (ghp_...):", "Set Token")
    If token = "" Then Exit Sub
    ThisWorkbook.Sheets("HistoryLog").Range("Z3").Value = token
    ThisWorkbook.Sheets("HistoryLog").Range("Z3").Font.Color = RGB(0, 0, 0)
    MsgBox "Token saved!", vbInformation
End Sub

Function FetchRawContent(filePath As String) As String
    Dim http As Object, ghToken As String
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ghToken = GetGitHubToken()
    Dim rawUrl As String
    rawUrl = "https://raw.githubusercontent.com/Timeverse/My-TW-Coverage/master/" & EncodeURL(filePath)
    On Error Resume Next
    With http
        .Open "GET", rawUrl, False
        .setRequestHeader "User-Agent", "VBA-Excel-Client"
        If ghToken <> "" Then .setRequestHeader "Authorization", "token " & ghToken
        .send
        If .status = 200 Then FetchRawContent = .responseText Else FetchRawContent = ""
    End With
    On Error GoTo 0
End Function

Sub RenderReportToSheet(content As String, tickerNum As String)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets("Coverage"): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Coverage"
    End If
    With ws
        .Cells.ClearContents: .Cells.ClearFormats
        .Cells.Interior.Color = RGB(10, 12, 18)
        .Cells.Font.Color = RGB(210, 210, 210)
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 100
        .Rows.RowHeight = 16
    End With

    Dim lines() As String: lines = Split(content, vbLf)
    Dim writeRow As Long: writeRow = 2
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim lt As String: lt = Replace(lines(i), vbCr, "")
        If lt = "" Then writeRow = writeRow + 1: GoTo NL
        With ws.Cells(writeRow, "B")
            Dim dt As String: dt = lt
            If Left(lt, 2) = "# " Then
                dt = CleanWikilinks(Mid(lt, 3))
                .Value = dt: .Font.Size = 18: .Font.Bold = True
                .Font.Color = RGB(255, 215, 0): .Interior.Color = RGB(25, 25, 35)
                writeRow = writeRow + 1
            ElseIf Left(lt, 3) = "## " Then
                dt = CleanWikilinks(Mid(lt, 4))
                .Value = "## " & dt: .Font.Size = 13: .Font.Bold = True
                .Font.Color = RGB(80, 180, 255): .Interior.Color = RGB(20, 22, 32)
            ElseIf Left(lt, 4) = "### " Then
                dt = CleanWikilinks(Mid(lt, 5))
                .Value = " ### " & dt: .Font.Size = 12: .Font.Bold = True
                .Font.Color = RGB(120, 230, 160)
            ElseIf Left(lt, 2) = "- " Or Left(lt, 2) = "* " Then
                .Value = " * " & CleanWikilinks(Mid(lt, 3))
                .Font.Color = RGB(200, 200, 200)
            ElseIf Left(lt, 1) = "|" Then
                .Value = CleanWikilinks(lt): .Font.Name = "Courier New"
                .Font.Size = 10: .Font.Color = RGB(180, 220, 180)
                .Interior.Color = RGB(18, 22, 28)
            ElseIf Left(lt, 3) = "```" Then
                GoTo NL
            Else
                .Value = RemoveBoldMarkers(CleanWikilinks(lt))
                .Font.Color = RGB(200, 200, 200)
            End If
        End With
        writeRow = writeRow + 1
NL:
    Next i
    ws.Activate: ws.Cells(1, 1).Select: ActiveWindow.ScrollRow = 1
End Sub

Function CleanWikilinks(text As String) As String
    Dim r As String: r = text
    Do While InStr(r, "[[") > 0
        Dim s As Long: s = InStr(r, "[[")
        Dim e As Long: e = InStr(s, r, "]]")
        If e > s Then
            r = Left(r, s - 1) & Mid(r, s + 2, e - s - 2) & Mid(r, e + 2)
        Else
            Exit Do
        End If
    Loop
    CleanWikilinks = r
End Function

Function RemoveBoldMarkers(text As String) As String
    RemoveBoldMarkers = Replace(text, "**", "")
End Function

' ================================================================
' Array Helpers
' ================================================================
Function CollectionToArray(col As Collection) As Variant
    If col.Count = 0 Then CollectionToArray = Empty: Exit Function
    Dim arr() As Variant: ReDim arr(0 To col.Count - 1)
    Dim i As Long
    For i = 0 To col.Count - 1: arr(i) = col(i + 1): Next i
    CollectionToArray = arr
End Function

Sub SortArray(ByRef arr As Variant)
    If IsEmpty(arr) Or Not IsArray(arr) Then Exit Sub
    If UBound(arr) - LBound(arr) < 1 Then Exit Sub
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
End Sub

' ================================================================
' DEBUG SYSTEM
' ================================================================
Sub RunSystemDebug()
    Application.ScreenUpdating = False

    Dim wsDB As Worksheet
    On Error Resume Next: Set wsDB = ThisWorkbook.Sheets("DebugLog"): On Error GoTo 0
    If wsDB Is Nothing Then
        Set wsDB = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDB.Name = "DebugLog"
    End If

    wsDB.Cells.Clear
    With wsDB.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
    wsDB.Activate
    ActiveWindow.DisplayGridlines = False
    wsDB.Columns("A").ColumnWidth = 22
    wsDB.Columns("B").ColumnWidth = 45
    wsDB.Columns("C").ColumnWidth = 18
    wsDB.Columns("D").ColumnWidth = 30

    With wsDB.Cells(1, 1)
        .Value = "SYSTEM DEBUG LOG | " & Format(Now, "yyyy/m/d hh:mm:ss")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 13
    End With
    wsDB.Rows(1).RowHeight = 26
    With wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(2, 4))
        .Interior.Color = RGB(255, 192, 0)
        .RowHeight = 3
    End With

    Dim wr As Long: wr = 4

    Call DB_Header(wsDB, wr, "CHECK", "ITEM", "STATUS", "DETAIL")
    wr = wr + 1

    ' 1. Worksheet Existence
    Call DB_Section(wsDB, wr, "WORKSHEET EXISTENCE")
    wr = wr + 1

    Dim sheetNames As Variant
    sheetNames = Array("RR4", "Transactions", "Realized", "HistoryLog", _
                       "Analysis", "HoldingsCorr", "Correlation", _
                       "DrawdownChart", "DebugLog")
    Dim sn As Variant
    For Each sn In sheetNames
        Dim wsCheck As Worksheet
        On Error Resume Next: Set wsCheck = ThisWorkbook.Sheets(CStr(sn)): On Error GoTo 0
        If wsCheck Is Nothing Then
            Call DB_Row(wsDB, wr, "Sheet", CStr(sn), "MISSING", "Sheet not found", RGB(255, 80, 80))
        Else
            Call DB_Row(wsDB, wr, "Sheet", CStr(sn), "OK", "Exists", RGB(0, 210, 100))
        End If
        wr = wr + 1
        Set wsCheck = Nothing
    Next sn

    ' 2. Transactions data check
    Call DB_Section(wsDB, wr, "TRANSACTIONS DATA")
    wr = wr + 1

    Dim wsTr As Worksheet
    On Error Resume Next: Set wsTr = ThisWorkbook.Sheets("Transactions"): On Error GoTo 0
    If Not wsTr Is Nothing Then
        Dim trLR As Long: trLR = wsTr.Cells(wsTr.Rows.Count, "A").End(xlUp).Row
        Dim trCount As Long: trCount = trLR - 1

        Call DB_Row(wsDB, wr, "Transactions", "Total rows", "INFO", trCount & " records", RGB(255, 192, 0))
        wr = wr + 1

        Dim blankDate As Long, blankTicker As Long, blankAction As Long, blankAmt As Long
        Dim tr As Long
        For tr = 2 To trLR
            If Trim(CStr(wsTr.Cells(tr, 2).Value)) = "" Then blankDate = blankDate + 1
            If Trim(CStr(wsTr.Cells(tr, 3).Value)) = "" Then blankTicker = blankTicker + 1
            If Trim(CStr(wsTr.Cells(tr, 4).Value)) = "" Then blankAction = blankAction + 1
            If Not IsNumeric(wsTr.Cells(tr, 9).Value) Then blankAmt = blankAmt + 1
        Next tr

        Call DB_Row(wsDB, wr, "Transactions", "Blank dates (Col B)", _
            IIf(blankDate = 0, "OK", "WARN"), blankDate & " rows", _
            IIf(blankDate = 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1
        Call DB_Row(wsDB, wr, "Transactions", "Blank Ticker (Col C)", _
            IIf(blankTicker = 0, "OK", "WARN"), blankTicker & " rows", _
            IIf(blankTicker = 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1
        Call DB_Row(wsDB, wr, "Transactions", "Blank Action (Col D)", _
            IIf(blankAction = 0, "OK", "WARN"), blankAction & " rows", _
            IIf(blankAction = 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1
        Call DB_Row(wsDB, wr, "Transactions", "Non-numeric NetAmt (Col I)", _
            IIf(blankAmt = 0, "OK", "WARN"), blankAmt & " rows", _
            IIf(blankAmt = 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1

        Dim invalidAction As Long
        For tr = 2 To trLR
            Dim act As String: act = UCase(Trim(CStr(wsTr.Cells(tr, 4).Value)))
            If act <> "BUY" And act <> "SELL" And act <> "ADJUSTCOST" And act <> "" Then
                invalidAction = invalidAction + 1
            End If
        Next tr
        Call DB_Row(wsDB, wr, "Transactions", "Invalid Action", _
            IIf(invalidAction = 0, "OK", "ERROR"), invalidAction & " rows (must be BUY/SELL/ADJUSTCOST)", _
            IIf(invalidAction = 0, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
    End If

    ' 3. HistoryLog update check
    Call DB_Section(wsDB, wr, "HISTORYLOG UPDATE CHECK")
    wr = wr + 1

    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = ThisWorkbook.Sheets("HistoryLog"): On Error GoTo 0
    If Not wsH Is Nothing Then
        Dim hlLR As Long: hlLR = wsH.Cells(wsH.Rows.Count, "A").End(xlUp).Row
        Dim hlCount As Long: hlCount = hlLR - 1
        Call DB_Row(wsDB, wr, "HistoryLog", "Total rows", "INFO", hlCount & " records", RGB(255, 192, 0))
        wr = wr + 1

        If hlLR >= 2 Then
            Dim lastLogTime As Variant: lastLogTime = wsH.Cells(hlLR, "A").Value
            Dim lastLogMkt As Variant: lastLogMkt = wsH.Cells(hlLR, "B").Value

            Dim timeDiff As Double
            If IsDate(lastLogTime) Then
                timeDiff = Now - CDate(lastLogTime)
                Dim diffMin As Long: diffMin = CLng(timeDiff * 1440)
                Dim timeStatus As String
                Dim timeColor As Long
                If diffMin < 5 Then
                    timeStatus = "FRESH": timeColor = RGB(0, 210, 100)
                ElseIf diffMin < 60 Then
                    timeStatus = "OK": timeColor = RGB(0, 210, 100)
                ElseIf diffMin < 480 Then
                    timeStatus = "STALE": timeColor = RGB(255, 140, 0)
                Else
                    timeStatus = "OLD": timeColor = RGB(255, 80, 80)
                End If
                Call DB_Row(wsDB, wr, "HistoryLog", "Last update time", timeStatus, _
                    Format(CDate(lastLogTime), "m/d hh:mm") & " (" & diffMin & " min ago)", timeColor)
            Else
                Call DB_Row(wsDB, wr, "HistoryLog", "Last update time", "ERROR", "Cannot parse date", RGB(255, 80, 80))
            End If
            wr = wr + 1

            If IsNumeric(lastLogMkt) And CDbl(lastLogMkt) > 0 Then
                Call DB_Row(wsDB, wr, "HistoryLog", "Last market value (Col B)", "OK", _
                    Format(CDbl(lastLogMkt), "#,##0") & " TWD", RGB(0, 210, 100))
            Else
                Call DB_Row(wsDB, wr, "HistoryLog", "Last market value (Col B)", "ERROR", "Zero or non-numeric", RGB(255, 80, 80))
            End If
            wr = wr + 1

            Dim mLR As Long: mLR = wsH.Cells(wsH.Rows.Count, "M").End(xlUp).Row
            Dim mCount As Long: mCount = mLR - 1
            Call DB_Row(wsDB, wr, "HistoryLog", "M-col data count (Port Ret)", _
                IIf(mCount > 0, "OK", "ERROR"), mCount & " rows", _
                IIf(mCount > 0, RGB(0, 210, 100), RGB(255, 80, 80)))
            wr = wr + 1

            Dim todayLogged As Boolean: todayLogged = False
            If IsDate(lastLogTime) Then
                If Int(CDate(lastLogTime)) = Date Then todayLogged = True
            End If
            Call DB_Row(wsDB, wr, "HistoryLog", "Logged today?", _
                IIf(todayLogged, "OK", "WARN"), _
                IIf(todayLogged, "Logged today", "Not yet logged today"), _
                IIf(todayLogged, RGB(0, 210, 100), RGB(255, 140, 0)))
            wr = wr + 1
        End If
    End If

    ' 4. RR4 Dashboard Check
    Call DB_Section(wsDB, wr, "RR4 DASHBOARD CHECK")
    wr = wr + 1

    Dim wsP As Worksheet
    On Error Resume Next: Set wsP = ThisWorkbook.Sheets("RR4"): On Error GoTo 0
    If Not wsP Is Nothing Then
        Dim exRate As Double: exRate = 0
        On Error Resume Next: exRate = CDbl(wsP.Range("C22").Value): On Error GoTo 0
        Call DB_Row(wsDB, wr, "RR4", "FX rate C22 (USD/TWD)", _
            IIf(exRate > 20 And exRate < 50, "OK", "WARN"), _
            IIf(exRate > 0, Format(exRate, "0.00"), "No value"), _
            IIf(exRate > 20 And exRate < 50, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1

        Dim totalMkt As Double: totalMkt = 0
        On Error Resume Next: totalMkt = CDbl(wsP.Cells(1, 1).Value): On Error GoTo 0
        Call DB_Row(wsDB, wr, "RR4", "Total market value (Cell A1)", _
            IIf(totalMkt > 0, "OK", "ERROR"), _
            IIf(totalMkt > 0, Format(totalMkt, "#,##0") & " TWD", "No value or zero"), _
            IIf(totalMkt > 0, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1

        Dim posLR As Long: posLR = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
        Dim posCount As Long: posCount = posLR - 4
        Call DB_Row(wsDB, wr, "RR4", "Position count (Row 5+)", _
            IIf(posCount > 0, "OK", "WARN"), posCount & " positions", _
            IIf(posCount > 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1

        Dim incDate As Date
        Dim stCap As Double
        On Error Resume Next
        incDate = CDate(wsP.Range("C2").Value)
        stCap = CDbl(wsP.Range("E2").Value)
        On Error GoTo 0
        Call DB_Row(wsDB, wr, "RR4", "InceptionDate (named range)", _
            IIf(incDate > DateSerial(2000, 1, 1), "OK", "ERROR"), _
            IIf(incDate > DateSerial(2000, 1, 1), Format(incDate, "yyyy/m/d"), "Not set"), _
            IIf(incDate > DateSerial(2000, 1, 1), RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
        Call DB_Row(wsDB, wr, "RR4", "StartingCapital (named range)", _
            IIf(stCap > 0, "OK", "ERROR"), _
            IIf(stCap > 0, Format(stCap, "#,##0") & " TWD", "Not set"), _
            IIf(stCap > 0, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
    End If

    ' 5. Yahoo Finance API Test
    Call DB_Section(wsDB, wr, "YAHOO FINANCE API TEST")
    wr = wr + 1

    Dim testTickers As Variant
    testTickers = Array("SPY", "QQQ", "2330.TW", "^TWII")
    Dim tt As Variant
    For Each tt In testTickers
        Dim httpT As Object
        On Error Resume Next
        Set httpT = CreateObject("WinHttp.WinHttpRequest.5.1")
        On Error GoTo 0

        Dim apiURL As String
        apiURL = "https://query1.finance.yahoo.com/v8/finance/chart/" & _
                 CStr(tt) & "?interval=1d&range=5d"

        Dim apiResp As String: apiResp = ""
        On Error Resume Next
        With httpT
            .Open "GET", apiURL, False
            .setRequestHeader "User-Agent", "Mozilla/5.0"
            .setRequestHeader "Accept", "application/json"
            .send
            apiResp = .responseText
        End With
        On Error GoTo 0

        Dim apiPx As Double: apiPx = 0
        Dim csPos As Long
        csPos = InStr(apiResp, """close"":[")
        If csPos > 0 Then
            Dim csS As Long: csS = csPos + 9
            Dim csE As Long: csE = InStr(csS, apiResp, "]")
            If csE > csS Then
                Dim csStr As String: csStr = Mid(apiResp, csS, csE - csS)
                Dim csParts() As String: csParts = Split(csStr, ",")
                Dim cpIdx As Long
                For cpIdx = UBound(csParts) To 0 Step -1
                    Dim cpTrim As String: cpTrim = Trim(csParts(cpIdx))
                    If IsNumeric(cpTrim) And CDbl(cpTrim) > 0 Then
                        apiPx = CDbl(cpTrim)
                        Exit For
                    End If
                Next cpIdx
            End If
        End If

        Dim httpStatus As Long: httpStatus = 0
        On Error Resume Next: httpStatus = httpT.status: On Error GoTo 0

        Dim detailMsg As String
        If apiPx > 0 Then
            detailMsg = "Price = " & Format(apiPx, "#,##0.00") & " (HTTP " & httpStatus & ")"
        ElseIf httpStatus = 429 Then
            detailMsg = "HTTP 429 - Too Many Requests (rate limited)"
        ElseIf httpStatus = 401 Or httpStatus = 403 Then
            detailMsg = "HTTP " & httpStatus & " - Auth required / forbidden"
        ElseIf httpStatus = 0 Then
            detailMsg = "No response (network or firewall)"
        Else
            detailMsg = "HTTP " & httpStatus & " - Cannot parse response"
        End If

        Call DB_Row(wsDB, wr, "API Test", CStr(tt), _
            IIf(apiPx > 0, "OK", "FAIL"), detailMsg, _
            IIf(apiPx > 0, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
    Next tt

    ' 6. Module function existence check
    Call DB_Section(wsDB, wr, "MODULE FUNCTION CHECK")
    wr = wr + 1

    Dim funcNames As Variant
    funcNames = Array("GetStockPrice", "GetCompanyName", "GetCurrencyType", _
                      "GetHistoryPrices", "GetTechData", "CorrPearson", _
                      "CorrGetCommonDates", "BuildPositions", "CalculateRealizedPnL")
    Dim fn As Variant
    For Each fn In funcNames
        Dim fnExists As Boolean: fnExists = False
        On Error Resume Next
        Dim testCall As Variant
        testCall = Application.Run(CStr(fn))
        If Err.Number = 0 Or Err.Number = 1004 Or Err.Number = 450 Or Err.Number = 13 Then
            fnExists = True
        ElseIf Err.Number = 1000 Or Err.Number = 35 Then
            fnExists = False
        Else
            fnExists = True
        End If
        Err.Clear
        On Error GoTo 0
        Call DB_Row(wsDB, wr, "Function", CStr(fn), _
            IIf(fnExists, "OK", "MISSING"), _
            IIf(fnExists, "Callable", "Function not found, check module"), _
            IIf(fnExists, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
    Next fn

    ' 7. Health Summary
    wr = wr + 1
    Call DB_Section(wsDB, wr, "HEALTH SUMMARY")
    wr = wr + 1

    Dim okCount As Long, warnCount As Long, errCount As Long
    Dim scanR As Long
    For scanR = 4 To wr
        Dim statusVal As String: statusVal = CStr(wsDB.Cells(scanR, 3).Value)
        Select Case statusVal
        Case "OK", "FRESH": okCount = okCount + 1
        Case "WARN", "STALE": warnCount = warnCount + 1
        Case "ERROR", "FAIL", "MISSING", "OLD": errCount = errCount + 1
        End Select
    Next scanR

    Dim healthScore As Long
    If okCount + warnCount + errCount > 0 Then
        healthScore = CLng(okCount / (okCount + warnCount + errCount) * 100)
    End If

    Dim scoreColor As Long
    If healthScore >= 90 Then
        scoreColor = RGB(0, 210, 100)
    ElseIf healthScore >= 70 Then
        scoreColor = RGB(255, 192, 0)
    Else
        scoreColor = RGB(255, 80, 80)
    End If

    With wsDB.Cells(wr, 1)
        .Value = "HEALTH SCORE"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    With wsDB.Cells(wr, 2)
        .Value = healthScore & " / 100"
        .Font.Color = scoreColor
        .Font.Bold = True
        .Font.Size = 14
    End With
    With wsDB.Cells(wr, 3)
        .Value = "OK:" & okCount & " WARN:" & warnCount & " ERR:" & errCount
        .Font.Color = RGB(200, 200, 200)
    End With

    wsDB.Columns("A:D").AutoFit
    wsDB.Activate
    wsDB.Cells(1, 1).Select

    Application.ScreenUpdating = True
    Application.StatusBar = "Debug complete - Health: " & healthScore & "/100"
    MsgBox "Debug complete. Health score: " & healthScore & "/100" & vbCr & _
           "OK: " & okCount & " WARN: " & warnCount & " ERROR: " & errCount, vbInformation
End Sub

' ----------------------------------------------------------------
Private Sub DB_Section(ws As Worksheet, r As Long, title As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Interior.Color = RGB(20, 20, 40)
        .Font.Color = RGB(100, 160, 255)
        .Font.Bold = True
    End With
    ws.Cells(r, 1).Value = "## " & title
    ws.Rows(r).RowHeight = 18
End Sub

' ----------------------------------------------------------------
Private Sub DB_Header(ws As Worksheet, r As Long, _
                      c1 As String, c2 As String, c3 As String, c4 As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Interior.Color = RGB(10, 10, 10)
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ws.Cells(r, 1).Value = c1
    ws.Cells(r, 2).Value = c2
    ws.Cells(r, 3).Value = c3
    ws.Cells(r, 4).Value = c4
    ws.Rows(r).RowHeight = 16
End Sub

' ----------------------------------------------------------------
Private Sub DB_Row(ws As Worksheet, r As Long, _
                   category As String, item As String, _
                   status As String, detail As String, statusColor As Long)
    ws.Rows(r).RowHeight = 16
    With ws.Cells(r, 1)
        .Value = category
        .Font.Color = RGB(150, 150, 150)
        .Interior.Color = IIf(r Mod 2 = 0, RGB(8, 8, 8), RGB(12, 12, 12))
    End With
    With ws.Cells(r, 2)
        .Value = item
        .Font.Color = RGB(200, 200, 200)
        .Interior.Color = IIf(r Mod 2 = 0, RGB(8, 8, 8), RGB(12, 12, 12))
    End With
    With ws.Cells(r, 3)
        .Value = status
        .Font.Color = statusColor
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = IIf(r Mod 2 = 0, RGB(8, 8, 8), RGB(12, 12, 12))
    End With
    With ws.Cells(r, 4)
        .Value = detail
        .Font.Color = RGB(180, 180, 180)
        .Interior.Color = IIf(r Mod 2 = 0, RGB(8, 8, 8), RGB(12, 12, 12))
    End With
End Sub
