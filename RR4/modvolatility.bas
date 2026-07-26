Attribute VB_Name = "modvolatility"
Option Explicit




'=============================================================================
'  功能：個股波動率分佈與進階風險分析 (支援多週期：100天, 200天, 1000天)
'  工作表：Tickers volatility
'  輸入：B2 (Ticker)
'  輸出：多週期報表 (起始列分別為 4, 29, 54)
'=============================================================================

Sub UpdateVolatilityAnalysis_Pro()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim allPrices() As Double, allDates() As Date
    Dim totalDays As Long
    
    ' 1. 設定工作表與輸入
    Set ws = ThisWorkbook.Sheets("Tickers volatility")
    Ticker = Trim(ws.Range("B2").Value)
    If Ticker = "" Then MsgBox "請在 B2 輸入股票代碼", vbExclamation: Exit Sub
    
    Application.ScreenUpdating = False
    
    ' 2. 清除舊數據 (範圍加大以容納三組分析表)
    ws.Range("A4:F100").ClearContents
    
    ' 3. 抓取歷史數據 (改為抓取 10 年，確保有 1000 個交易日)
    If Not GetHistoricalData(Ticker, allDates, allPrices) Then
        MsgBox "無法抓取數據，請檢查代碼或網絡。", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    totalDays = UBound(allPrices) + 1
    If totalDays < 20 Then
        MsgBox "數據筆數不足以進行進階分析。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 4. 執行三個不同週期的分析 (傳入: 工作表, 全部價格, 欲分析天數, 輸出起始列, 標題)
    ' 1-a. 過去 100 天 (輸出在 A4:F26)
    Call AnalyzeAndOutput(ws, allPrices, 100, 4, "Past 100 Days")
    
    ' 1-b. 過去 200 天 (輸出在 A29:F51)
    Call AnalyzeAndOutput(ws, allPrices, 200, 29, "Past 200 Days")
    
    ' 1-c. 過去 1000 天 (輸出在 A54:F76)
    Call AnalyzeAndOutput(ws, allPrices, 1000, 54, "Past 1000 Days")
    
    Application.ScreenUpdating = True
    MsgBox "100天、200天、1000天進階波動率分析更新完成！", vbInformation
End Sub

'=============================================================================
'  副程式：執行特定天數的計算與報表輸出
'=============================================================================
Sub AnalyzeAndOutput(ws As Worksheet, ByRef allPrices() As Double, lookbackDays As Long, startRow As Long, titleSuffix As String)
    Dim prices() As Double, returns() As Double
    Dim i As Long, j As Long, n As Long
    Dim totalData As Long, useDays As Long
    
    totalData = UBound(allPrices) + 1
    useDays = lookbackDays
    If totalData < useDays Then useDays = totalData ' 如果歷史數據不夠，則用最大可用的天數
    
    ' 1. 截取最近 useDays 的股價數據
    ReDim prices(0 To useDays - 1)
    For i = 0 To useDays - 1
        prices(i) = allPrices(UBound(allPrices) - useDays + 1 + i)
    Next i
    
    n = UBound(prices)
    If n < 1 Then Exit Sub
    ReDim returns(0 To n - 1)
    
    ' --- 統計變數宣告 ---
    Dim mean As Double, stdDev As Double, maxRet As Double, minRet As Double
    Dim annVol As Double, skew As Double, kurt As Double
    Dim downDev As Double, mdd As Double
    Dim currRollVol As Double, maxRollVol As Double
    Dim upStreak As Long, downStreak As Long, maxUpStreak As Long, maxDownStreak As Long
    
    ' 2. 計算日漲跌幅 & 連續漲跌天數
    maxUpStreak = 0: maxDownStreak = 0
    upStreak = 0: downStreak = 0
    
    For i = 1 To n
        If prices(i - 1) <> 0 Then
            returns(i - 1) = (prices(i) - prices(i - 1)) / prices(i - 1)
        Else
            returns(i - 1) = 0
        End If
        
        If returns(i - 1) > 0 Then
            upStreak = upStreak + 1
            downStreak = 0
            If upStreak > maxUpStreak Then maxUpStreak = upStreak
        ElseIf returns(i - 1) < 0 Then
            downStreak = downStreak + 1
            upStreak = 0
            If downStreak > maxDownStreak Then maxDownStreak = downStreak
        Else
            upStreak = 0
            downStreak = 0
        End If
    Next i
    
    ' 3. 計算最大回撤 (MDD)
    Dim peak As Double, DrawDown As Double
    peak = prices(0)
    mdd = 0
    For i = 1 To n
        If prices(i) > peak Then peak = prices(i)
        DrawDown = (peak - prices(i)) / peak
        If DrawDown > mdd Then mdd = DrawDown
    Next i
    
    ' 4. 計算下行標準差 (Downside Deviation)
    Dim downReturns() As Double
    Dim downCount As Long
    downCount = 0
    For i = LBound(returns) To UBound(returns)
        If returns(i) < 0 Then
            ReDim Preserve downReturns(0 To downCount)
            downReturns(downCount) = returns(i)
            downCount = downCount + 1
        End If
    Next i
    
    ' 5. 計算 20日滾動波動率
    Dim rollWindow(1 To 20) As Double
    Dim rollVol As Double
    maxRollVol = 0
    If UBound(returns) >= 19 Then
        For i = 19 To UBound(returns)
            For j = 1 To 20
                rollWindow(j) = returns(i - 20 + j)
            Next j
            rollVol = Application.WorksheetFunction.StDev_S(rollWindow) * Sqr(252)
            If rollVol > maxRollVol Then maxRollVol = rollVol
            If i = UBound(returns) Then currRollVol = rollVol
        Next i
    End If
    
    ' 6. 基礎與進階統計指標
    On Error Resume Next
    mean = Application.WorksheetFunction.Average(returns)
    stdDev = Application.WorksheetFunction.StDev_S(returns)
    maxRet = Application.WorksheetFunction.Max(returns)
    minRet = Application.WorksheetFunction.Min(returns)
    annVol = stdDev * Sqr(252)
    skew = Application.WorksheetFunction.skew(returns)
    kurt = Application.WorksheetFunction.kurt(returns)
    If downCount > 1 Then
        downDev = Application.WorksheetFunction.StDev_S(downReturns)
    Else
        downDev = 0
    End If
    On Error GoTo 0
    
    ' 7. 填寫標題與數值 (結構化輸出)
    Dim labelsArr As Variant, valsArr As Variant
    labelsArr = Array("Mean (Daily)", "Std Dev (Daily)", "Annualized Volatility", _
                      "Skewness (不對稱性)", "Kurtosis (肥尾效應)", _
                      "Downside Deviation", "Max Drawdown (MDD)", _
                      "Current 20D Volatility", "Max 20D Volatility", _
                      "Max Up Streak (Days)", "Max Down Streak (Days)", _
                      "Max Daily Chg%", "Min Daily Chg%")
                      
    valsArr = Array(mean, stdDev, annVol, skew, kurt, downDev, -mdd, _
                    currRollVol, maxRollVol, maxUpStreak, maxDownStreak, maxRet, minRet)
    
    ' 寫入表頭
    ws.cells(startRow, 1).Value = "Statistic (" & titleSuffix & ")"
    ws.cells(startRow, 2).Value = "Value"
    
    For i = LBound(labelsArr) To UBound(labelsArr)
        ws.cells(startRow + 1 + i, 1).Value = labelsArr(i)
        ws.cells(startRow + 1 + i, 2).Value = valsArr(i)
    Next i
    
    ' 動態套用格式化
    ws.Range("B" & startRow + 1 & ":B" & startRow + 3 & ", B" & startRow + 6 & ":B" & startRow + 9 & ", B" & startRow + 12 & ":B" & startRow + 13).NumberFormat = "0.00%"
    ws.Range("B" & startRow + 4 & ":B" & startRow + 5).NumberFormat = "0.00"
    ws.Range("B" & startRow + 10 & ":B" & startRow + 11).NumberFormat = "0"
    
    ' 8. 建立分佈表
    Dim bins(0 To 21) As Long, labels(0 To 21) As String
    Dim retVal As Double, binIdx As Integer
    Dim totalReturnsCount As Long
    
    totalReturnsCount = UBound(returns) + 1
    labels(0) = "< -10%"
    labels(21) = "> 10%"
    
    For i = 1 To 20
        Dim lowerLimit As Double
        lowerLimit = -0.1 + (i - 1) * 0.01
        labels(i) = Format(lowerLimit, "0%") & " ~ " & Format(lowerLimit + 0.01, "0%")
    Next i
    
    For i = LBound(returns) To UBound(returns)
        retVal = returns(i)
        If retVal < -0.1 Then
            bins(0) = bins(0) + 1
        ElseIf retVal >= 0.1 Then
            bins(21) = bins(21) + 1
        Else
            binIdx = Int((retVal + 0.1) / 0.01) + 1
            If binIdx >= 1 And binIdx <= 20 Then bins(binIdx) = bins(binIdx) + 1
        End If
    Next i
    
    ' 填入分佈表格
    ws.cells(startRow, 4).Value = "Range"
    ws.cells(startRow, 5).Value = "Count"
    ws.cells(startRow, 6).Value = "Freq %"
    
    For i = 0 To 21
        ws.cells(startRow + 1 + i, 4).Value = "'" & labels(i)
        ws.cells(startRow + 1 + i, 5).Value = bins(i)
        If totalReturnsCount > 0 Then
            ws.cells(startRow + 1 + i, 6).Value = bins(i) / totalReturnsCount
        Else
            ws.cells(startRow + 1 + i, 6).Value = 0
        End If
    Next i
    ws.Range("F" & startRow + 1 & ":F" & startRow + 22).NumberFormat = "0.00%"

End Sub

'=============================================================================
'  核心函數：抓取 10 年數據 (供切割週期使用)
'=============================================================================
Function GetHistoricalData(Ticker As String, ByRef outDates() As Date, ByRef outPrices() As Double) As Boolean
    Dim http As Object
    Dim url As String, response As String
    Dim tsStr As String, closeStr As String
    Dim tsArr() As String, closeArr() As String
    Dim i As Long
    
    ' 1. 處理 Ticker
    If InStr(Ticker, ".TW") = 0 And InStr(Ticker, ".TWO") = 0 And IsNumeric(Ticker) Then
        Ticker = Ticker & ".TW"
    End If
    
    ' 2. 設定 URL (Range 改為 10y 確保涵蓋 1000 交易日)
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?range=10y&interval=1d"
    
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        response = .responseText
    End With
    On Error GoTo 0
    
    ' 3. 解析 JSON
    Dim tsStart As Long, tsEnd As Long
    tsStart = InStr(response, """timestamp"":[") + 12
    tsEnd = InStr(tsStart, response, "]")
    If tsStart > 12 And tsEnd > tsStart Then
        tsStr = Mid(response, tsStart, tsEnd - tsStart)
        tsArr = Split(tsStr, ",")
    Else
        GetHistoricalData = False
        Exit Function
    End If
    
    Dim closeStart As Long, closeEnd As Long
    closeStart = InStr(response, """adjclose"":[{""adjclose"":[") + 23
    If closeStart < 24 Then closeStart = InStr(response, """close"":[") + 8
    
    closeEnd = InStr(closeStart, response, "]")
    If closeStart > 20 And closeEnd > closeStart Then
        closeStr = Mid(response, closeStart, closeEnd - closeStart)
        closeArr = Split(closeStr, ",")
    Else
        GetHistoricalData = False
        Exit Function
    End If
    
    Dim count As Long
    count = UBound(tsArr)
    If UBound(closeArr) < count Then count = UBound(closeArr)
    
    ReDim outDates(0 To count)
    ReDim outPrices(0 To count)
    
    Dim validCount As Long
    validCount = 0
    
    For i = 0 To count
        If IsNumeric(tsArr(i)) And IsNumeric(closeArr(i)) Then
            outDates(validCount) = (CDbl(tsArr(i)) / 86400) + 25569
            outPrices(validCount) = CDbl(closeArr(i))
            validCount = validCount + 1
        End If
    Next i
    
    If validCount > 0 Then
        ReDim Preserve outDates(0 To validCount - 1)
        ReDim Preserve outPrices(0 To validCount - 1)
        GetHistoricalData = True
    Else
        GetHistoricalData = False
    End If
End Function

