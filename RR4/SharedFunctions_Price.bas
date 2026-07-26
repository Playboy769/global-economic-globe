Attribute VB_Name = "SharedFunctions_Price"

Option Explicit

' ================================================================
'  FetchTWSE
'  ㊣ TWSE/TPEX  API眔程穝Θユ基絃い┪Μ絃基
'  ex_ch 絛ㄒ: "tse_2337.tw" ┪ "otc_6510.tw"
' ================================================================
Private Function FetchTWSE(http As Object, exch As String) As Double

    Dim url As String
    url = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=" & _
          exch & "&json=1&delay=0&_=" & CStr(CLng((Now - DateSerial(1970, 1, 1)) * 86400))

    Dim response As String
    On Error Resume Next
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .setRequestHeader "Referer", "https://mis.twse.com.tw/"
        .setRequestHeader "Cache-Control", "no-cache"
        .send
        If .status = 200 Then response = .responseText
    End With
    On Error GoTo 0

    If Len(response) < 20 Then FetchTWSE = 0: Exit Function

    '  纔 "z"程穝Θユ基絃い/Μ絃
    Dim price As Double
    price = TWSEExtract(response, """z"":""")

    '  璝 z  "-"ヰカ/ゼ秨絃э琎らΜ絃 "y" 
    If price = 0 Then price = TWSEExtract(response, """y"":""")

    FetchTWSE = price
End Function

' ================================================================
'  TWSEExtract  眖 TWSE JSON 计逆
'  Α  "z":"120.50"  ┪  "y":"121.00"
' ================================================================
Private Function TWSEExtract(json As String, tag As String) As Double
    Dim pos As Long: pos = InStr(json, tag)
    If pos = 0 Then TWSEExtract = 0: Exit Function

    pos = pos + Len(tag)
    Dim ep As Long: ep = pos
    Do While ep <= Len(json)
        Select Case Mid(json, ep, 1)
            Case """", ",", "}", "]": Exit Do
        End Select
        ep = ep + 1
    Loop

    Dim s As String: s = Trim(Mid(json, pos, ep - pos))
    If s = "-" Or s = "" Then TWSEExtract = 0: Exit Function
    If IsNumeric(s) Then TWSEExtract = CDbl(s)
End Function

' ================================================================
'  FetchYahooClose  ㄏノ interval=1d&range=5d
'  眖 close[] 皚程Τタ计
' ================================================================
Private Function FetchYahooClose(http As Object, Ticker As String) As Double

    Dim url As String
    url = "https://query2.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=5d"

    Dim response As String
    On Error Resume Next
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        .setRequestHeader "Referer", "https://finance.yahoo.com/"
        .setRequestHeader "Cache-Control", "no-cache"
        .send
        If .status = 200 Then response = .responseText
    End With
    On Error GoTo 0

    If Len(response) < 200 Then FetchYahooClose = 0: Exit Function

    ' ﹚ quote 跋遏ず close[]
    Dim qPos As Long: qPos = InStr(response, """quote"":{")
    If qPos = 0 Then qPos = 1

    Dim csStart As Long: csStart = InStr(qPos, response, """close"":[")
    If csStart = 0 Then FetchYahooClose = 0: Exit Function
    csStart = csStart + 9

    Dim csEnd As Long: csEnd = InStr(csStart, response, "]")
    If csEnd <= csStart Then FetchYahooClose = 0: Exit Function

    Dim closeArr() As String
    closeArr = Split(Mid(response, csStart, csEnd - csStart), ",")

    Dim i As Long, price As Double
    For i = UBound(closeArr) To LBound(closeArr) Step -1
        Dim cv As String: cv = Trim(closeArr(i))
        ' タ絋糶猭╊Θㄢ糷 If
    If IsNumeric(cv) Then
        If CDbl(cv) > 0 Then
            price = CDbl(cv)
            Exit For
        End If
    End If
    Next i

    FetchYahooClose = price
End Function

' ================================================================
'  DebugPriceTest  v6.0
' ================================================================
Sub DebugPriceTest()

    Dim testTicker As String
    testTicker = InputBox("块布腹 2337 ┪ AAPL:", "Price Debug v6", "2337")
    If testTicker = "" Then Exit Sub

    testTicker = UCase(Trim(testTicker))
    If IsNumeric(testTicker) Then testTicker = testTicker & ".TW"
    If Not g_PriceCache Is Nothing Then g_PriceCache.RemoveAll

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    Dim isTW As Boolean
    isTW = (InStr(testTicker, ".TW") > 0 Or InStr(testTicker, ".TWO") > 0)

    Dim twseResp As String, yahooResp As String
    Dim TWSEPrice As Double, yahooPrice As Double, finalPrice As Double
    Dim code As String

    If isTW Then
        code = Replace(Replace(testTicker, ".TWO", ""), ".TW", "")

        ' TWSE 叫―
        Dim twseUrl As String
        twseUrl = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_" & _
                  code & ".tw&json=1&delay=0"
        With http
            .Open "GET", twseUrl, False
            .setRequestHeader "User-Agent", "Mozilla/5.0"
            .setRequestHeader "Referer", "https://mis.twse.com.tw/"
            .send
            twseResp = .responseText
        End With
        TWSEPrice = FetchTWSE(http, "tse_" & code & ".tw")
    End If

    ' Yahoo 叫―
    Dim yahooUrl As String
    yahooUrl = "https://query2.finance.yahoo.com/v8/finance/chart/" & testTicker & "?interval=1d&range=5d"
    With http
        .Open "GET", yahooUrl, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .setRequestHeader "Cache-Control", "no-cache"
        .send
        yahooResp = .responseText
    End With
    yahooPrice = FetchYahooClose(http, testTicker)

    finalPrice = GetStockPrice(testTicker)

    Dim info As String
    If isTW Then
        info = "TWSE API" & vbCr & _
               "URL    : tse_" & code & ".tw" & vbCr & _
               "基   : " & TWSEPrice & vbCr & _
               "﹍   : " & Left(twseResp, 300) & vbCr & vbCr
    End If
    info = info & "Yahoo Finance" & vbCr & _
           "URL    : " & yahooUrl & vbCr & _
           "Len    : " & Len(yahooResp) & " chars" & vbCr & _
           "基   : " & yahooPrice & vbCr & vbCr & _
           "」 GetStockPrice 程沧挡狦: " & finalPrice

    MsgBox info, vbInformation, "Price Debug v6: " & testTicker
    Set http = Nothing
End Sub


