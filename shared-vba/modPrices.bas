Attribute VB_Name = "modPrices"
Option Explicit

' Stock price history. SEC EDGAR only has fundamentals, not market prices, so this
' hits Yahoo Finance's public chart JSON endpoint (no API key required) for
' US-listed tickers. (Stooq's CSV endpoint now sits behind a JS proof-of-work
' challenge that a plain HTTP client can't solve, so that's not an option here.)
Public Function FetchDailyPrices(ByVal ticker As String, ByVal fromDate As Date, ByVal toDate As Date) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim epoch As Date
    epoch = DateSerial(1970, 1, 1)
    Dim p1 As Double, p2 As Double
    p1 = DateDiff("s", epoch, fromDate)
    p2 = DateDiff("s", epoch, toDate) + 86400

    Dim url As String
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & ticker & "?period1=" & CLng(p1) & "&period2=" & CLng(p2) & "&interval=1d"

    Dim json As String
    On Error Resume Next
    json = HttpGet(url)
    On Error GoTo 0
    If Len(json) = 0 Then
        Set FetchDailyPrices = dict
        Exit Function
    End If

    Dim chartRaw As String
    chartRaw = ExtractJsonValueRaw(json, "chart")
    Dim resultArrRaw As String
    resultArrRaw = ExtractJsonValueRaw(chartRaw, "result")
    Dim resultElems As Collection
    Set resultElems = SplitJsonArrayElements(resultArrRaw)
    If resultElems.Count = 0 Then
        Set FetchDailyPrices = dict
        Exit Function
    End If

    Dim resultObj As String
    resultObj = CStr(resultElems(1))

    Dim timestampArrRaw As String
    timestampArrRaw = ExtractJsonValueRaw(resultObj, "timestamp")
    Dim indicatorsRaw As String
    indicatorsRaw = ExtractJsonValueRaw(resultObj, "indicators")
    Dim quoteArrRaw As String
    quoteArrRaw = ExtractJsonValueRaw(indicatorsRaw, "quote")
    Dim quoteElems As Collection
    Set quoteElems = SplitJsonArrayElements(quoteArrRaw)
    If quoteElems.Count = 0 Then
        Set FetchDailyPrices = dict
        Exit Function
    End If
    Dim closeArrRaw As String
    closeArrRaw = ExtractJsonValueRaw(CStr(quoteElems(1)), "close")

    Dim timestamps As Collection, closes As Collection
    Set timestamps = SplitJsonArrayElements(timestampArrRaw)
    Set closes = SplitJsonArrayElements(closeArrRaw)

    Dim n As Long
    n = timestamps.Count
    If closes.Count < n Then n = closes.Count

    Dim i As Long
    For i = 1 To n
        Dim tsStr As String, clStr As String
        tsStr = CStr(timestamps(i))
        clStr = CStr(closes(i))
        If IsNumeric(tsStr) And IsNumeric(clStr) Then
            Dim dt As Date
            dt = DateAdd("s", CDbl(tsStr), epoch)
            dict(Format$(dt, "yyyy-mm-dd")) = CDbl(clStr)
        End If
    Next i

    Set FetchDailyPrices = dict
End Function

' Most recent available close (any of the last ~10 calendar days), for
' contexts that just want "the current price" cheaply -- e.g. Watchlist row
' refreshes, where fetching a multi-year history per ticker would be wasteful.
' Returns "" if nothing came back (bad ticker, market holiday streak, etc.).
Public Function LatestPrice(ByVal ticker As String) As Variant
    Dim prices As Object
    Set prices = FetchDailyPrices(ticker, DateAdd("d", -10, Date), Date)
    If prices.Count = 0 Then
        LatestPrice = ""
        Exit Function
    End If

    Dim bestKey As String
    Dim k As Variant
    For Each k In prices.Keys
        If bestKey = "" Or CStr(k) > bestKey Then bestKey = CStr(k)
    Next k
    LatestPrice = prices(bestKey)
End Function

' Latest available close on or before targetDateStr ("yyyy-mm-dd"). Returns "" if
' no price exists at or before that date (e.g. IPO'd after that year).
Public Function NearestPriceOnOrBefore(ByVal prices As Object, ByVal targetDateStr As String) As Variant
    Dim bestKey As String
    Dim bestVal As Variant
    bestVal = ""
    Dim k As Variant
    For Each k In prices.Keys
        If CStr(k) <= targetDateStr Then
            If bestKey = "" Or CStr(k) > bestKey Then
                bestKey = CStr(k)
                bestVal = prices(k)
            End If
        End If
    Next k
    NearestPriceOnOrBefore = bestVal
End Function
