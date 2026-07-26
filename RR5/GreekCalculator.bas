Attribute VB_Name = "GreekCalculator"
Option Explicit

' =========================================================
' 1. 初始化計算機介面
' =========================================================
Sub SetupCalculatorGUI()
    Dim wS As Worksheet
    Set wS = ActiveSheet

    ' 清除現有工作表內容與格式
    wS.Cells.Clear

    ' 標題設定
    wS.Range("B2").Value = "選擇權單次計算機 (Options Calculator)"
    wS.Range("B2").Font.Bold = True
    wS.Range("B2").Font.Size = 12
    wS.Columns("B").ColumnWidth = 24
    wS.Columns("C").ColumnWidth = 15

    ' --- 建立輸入區標籤 ---
    wS.Range("B4").Value = "評價日期 (Date):"
    wS.Range("B5").Value = "標的代碼 (Ticker):"
    wS.Range("B6").Value = "選擇權類型 (Type):"
    wS.Range("B7").Value = "履約價 (Strike):"
    wS.Range("B8").Value = "到期日 (Expiry):"
    wS.Range("B9").Value = "標的現價 (Underlying):"
    wS.Range("B10").Value = "進場 IV (Entry IV):"
    wS.Range("B11").Value = "無風險利率 (Risk-Free Rate):"
    wS.Range("B12").Value = "行使比例 (Exercise Ratio):" ' [新增] 行使比例欄位

    ' --- 建立輸出區標籤 ---
    wS.Range("B14").Value = "[ 系統計算結果 ]"
    wS.Range("B14").Font.Bold = True
    wS.Range("B15").Value = "Gamma (Γ):"
    wS.Range("B16").Value = "每日 Theta (Θ):"

    ' --- 視覺化格式設定 ---
    ' 將輸入區 (C4:C12) 設為淺黃色，加上外框
    Dim inputRange As Range
    Set inputRange = wS.Range("C4:C12")
    inputRange.Interior.Color = RGB(255, 255, 224)
    inputRange.Borders.LineStyle = xlContinuous

    ' 將輸出區 (C15:C16) 設為淺綠色，加上外框
    Dim outputRange As Range
    Set outputRange = wS.Range("C15:C16")
    outputRange.Interior.Color = RGB(224, 255, 224)
    outputRange.Borders.LineStyle = xlContinuous
    outputRange.NumberFormat = "0.0000"

    ' --- 填入預設測試資料 ---
    wS.Range("C4").Value = Date
    wS.Range("C5").Value = "2330.TW"
    wS.Range("C6").Value = "Call"
    wS.Range("C7").Value = 2400
    wS.Range("C8").Value = Date + 30
    wS.Range("C9").Value = 2300
    wS.Range("C10").Value = 0.45
    wS.Range("C11").Value = 0.02
    wS.Range("C12").Value = 0.1 ' [新增] 預設行使比例 (如 0.1)

    MsgBox "計算機面板已建置完成。" & vbCrLf & _
           "已新增「行使比例」欄位。請填妥參數後，執行 CalculateGreeks 進行運算。", vbInformation, "初始化成功"
End Sub

' =========================================================
' 2. 執行 Greeks 計算
' =========================================================
Sub CalculateGreeks()
    Dim wS As Worksheet
    Set wS = ActiveSheet
    
    ' 宣告計算變數
    Dim S As Double, k As Double, T As Double, v As Double, r As Double
    Dim exerciseRatio As Double
    Dim optType As String
    Dim calcGamma As Double
    Dim calcTheta As Double
    
    On Error GoTo ErrorHandler
    
    ' 從介面讀取變數
    S = CDbl(wS.Range("C9").Value)
    k = CDbl(wS.Range("C7").Value)
    v = CDbl(wS.Range("C10").Value)
    optType = CStr(wS.Range("C6").Value)
    
    ' 動態讀取無風險利率 (C11)
    If Trim(wS.Range("C11").Value) = "" Then
        r = 0.02
    Else
        r = CDbl(wS.Range("C11").Value)
        If r > 0.5 Then r = r / 100
    End If
    
    ' [新增] 動態讀取行使比例 (C12)
    If Trim(wS.Range("C12").Value) = "" Then
        exerciseRatio = 1 ' 若空白則預設為 1 (標準選擇權)
    Else
        exerciseRatio = CDbl(wS.Range("C12").Value)
    End If
    
    ' 計算距離到期年限 T
    T = (CDate(wS.Range("C8").Value) - CDate(wS.Range("C4").Value)) / 365
    
    If T <= 0 Then
        MsgBox "運算中斷：到期日 (Expiry) 必須晚於評價日期 (Date)。", vbExclamation, "邏輯錯誤"
        Exit Sub
    End If

    ' 呼叫 Math_Greeks 函數計算，並乘上行使比例
    calcGamma = Calc_Gamma(S, k, T, r, v) * exerciseRatio
    calcTheta = Calc_Daily_Theta(S, k, T, r, v, optType) * exerciseRatio
    
    ' 將結果輸出至介面 (位置已配合 UI 下移至 15, 16 列)
    wS.Range("C15").Value = calcGamma
    wS.Range("C16").Value = calcTheta
    
    Exit Sub

ErrorHandler:
    MsgBox "計算中斷。請確認黃色輸入欄位皆包含正確的數值與日期格式。", vbCritical, "資料讀取失敗"
End Sub

' =========================================================
' 3. 執行 Theta 時間旅行矩陣
' =========================================================
Sub GenerateThetaData()
    Dim wS As Worksheet
    Set wS = ActiveSheet
    
    Dim S As Double, k As Double, v As Double, r As Double
    Dim exerciseRatio As Double
    Dim optType As String
    Dim totalDays As Integer
    
    On Error GoTo ErrorHandler
    
    S = CDbl(wS.Range("C9").Value)
    k = CDbl(wS.Range("C7").Value)
    v = CDbl(wS.Range("C10").Value)
    optType = CStr(wS.Range("C6").Value)
    
    If Trim(wS.Range("C11").Value) = "" Then
        r = 0.02
    Else
        r = CDbl(wS.Range("C11").Value)
        If r > 0.5 Then r = r / 100
    End If
    
    ' [新增] 動態讀取行使比例
    If Trim(wS.Range("C12").Value) = "" Then
        exerciseRatio = 1
    Else
        exerciseRatio = CDbl(wS.Range("C12").Value)
    End If
    
    If v > 2 Then v = v / 100
    
    totalDays = CInt(CDate(wS.Range("C8").Value) - CDate(wS.Range("C4").Value))
    If totalDays <= 0 Then
        MsgBox "運算中斷：到期日必須晚於評價日。", vbExclamation
        Exit Sub
    End If
    
    ' [修正] 將資料輸出區起始列下移至第 19 列，避免覆蓋計算機結果
    Dim dataStartRow As Long, i As Long
    Dim currentT As Double, dailyTheta As Double
    dataStartRow = 19
    
    wS.Range("B" & dataStartRow & ":C" & wS.Rows.Count).ClearContents
    wS.Range("B" & dataStartRow & ":C" & wS.Rows.Count).Font.Color = vbBlack ' 修改為黑色確保可見
    
    wS.Cells(dataStartRow, 2).Value = "Days to Expiration"
    wS.Cells(dataStartRow, 3).Value = "Daily Theta (Warrant)"
    wS.Cells(dataStartRow, 2).Font.Bold = True
    wS.Cells(dataStartRow, 3).Font.Bold = True
    
    Dim rowOffset As Long
    rowOffset = 1
    
    For i = totalDays To 1 Step -1
        currentT = i / 365
        
        ' 算出的單日 Theta 乘上行使比例
        dailyTheta = Calc_Daily_Theta(S, k, currentT, r, v, optType) * exerciseRatio
        
        wS.Cells(dataStartRow + rowOffset, 2).Value = i
        wS.Cells(dataStartRow + rowOffset, 3).Value = dailyTheta
        
        rowOffset = rowOffset + 1
    Next i
    
    wS.Range(wS.Cells(dataStartRow + 1, 3), wS.Cells(dataStartRow + totalDays, 3)).NumberFormat = "0.0000"
    
    MsgBox "Theta 數據已成功產出！", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "產生資料失敗，請檢查輸入參數是否完整無誤。", vbCritical
End Sub

' =========================================================
' 核心數學函數 (Math_Greeks)
' =========================================================
Function Calc_d1(S As Double, k As Double, T As Double, r As Double, v As Double) As Double
    Calc_d1 = (Log(S / k) + (r + v ^ 2 / 2) * T) / (v * Sqr(T))
End Function

Function ND_prime(x As Double) As Double
    ND_prime = Exp(-(x ^ 2) / 2) / Sqr(2 * WorksheetFunction.pi())
End Function

Function Calc_Gamma(S As Double, k As Double, T As Double, r As Double, v As Double) As Double
    Dim d1 As Double
    If T <= 0 Then T = 0.0001
    d1 = Calc_d1(S, k, T, r, v)
    Calc_Gamma = ND_prime(d1) / (S * v * Sqr(T))
End Function

Function Calc_Daily_Theta(S As Double, k As Double, T As Double, r As Double, v As Double, optType As String) As Double
    Dim d1 As Double, d2 As Double
    Dim Nd1p As Double, Nd2 As Double, N_minus_d2 As Double
    Dim annualTheta As Double
    
    If T <= 0 Then T = 0.0001
    
    d1 = Calc_d1(S, k, T, r, v)
    d2 = d1 - v * Sqr(T)
    Nd1p = ND_prime(d1)
    
    If UCase(optType) = "CALL" Or UCase(optType) = "C" Then
        Nd2 = WorksheetFunction.Norm_Dist(d2, 0, 1, True)
        annualTheta = -(S * v * Nd1p) / (2 * Sqr(T)) - r * k * Exp(-r * T) * Nd2
    ElseIf UCase(optType) = "PUT" Or UCase(optType) = "P" Then
        N_minus_d2 = WorksheetFunction.Norm_Dist(-d2, 0, 1, True)
        annualTheta = -(S * v * Nd1p) / (2 * Sqr(T)) + r * k * Exp(-r * T) * N_minus_d2
    End If
    
    Calc_Daily_Theta = annualTheta / 365
End Function

Function GetLivePrice(ticker As String) As Double
    Dim req As Object
    Dim url As String
    Dim response As String
    Dim priceStart As Long
    Dim priceEnd As Long
    Dim priceStr As String
    
    On Error GoTo ErrorHandler
    
    Set req = CreateObject("MSXML2.XMLHTTP")
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & ticker
    
    req.Open "GET", url, False
    req.send
    response = req.responseText
    
    priceStart = InStr(response, """regularMarketPrice"":") + 21
    If priceStart > 21 Then
        priceEnd = InStr(priceStart, response, ",")
        priceStr = Mid(response, priceStart, priceEnd - priceStart)
        GetLivePrice = CDbl(priceStr)
    Else
        GetLivePrice = 0
    End If
    Exit Function

ErrorHandler:
    GetLivePrice = 0
End Function

