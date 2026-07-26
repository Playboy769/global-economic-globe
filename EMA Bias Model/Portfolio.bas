Attribute VB_Name = "Portfolio"
Option Explicit

' ============================================================
'  Portfolio Builder (依附 EMA_Bias_Model 模組)
'  - Portfolio 分頁 A 欄輸入 ticker 清單
'  - BuildPortfolio: 逐檔跑模型，抽矩陣寫到 PortfolioReturns / RawReturns / Positions
'  - BuildAllModels: 為每檔 ticker 建立獨立 Model 分頁 (M_<ticker>)
'  - DeleteAllModelSheets: 一鍵清除所有 M_<ticker> 分頁
'  - ExportAllCSV: 匯出 3 個 CSV 給 Python
' ============================================================

' Ticker → 合法 Excel sheet 名稱 (M_<ticker>，去除 : \ / ? * [ ] 與長度限制)
Function ModelSheetName(ticker As String) As String
    Dim s As String
    s = "M_" & ticker
    s = Replace(s, ":", "_")
    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "[", "_")
    s = Replace(s, "]", "_")
    If Len(s) > 31 Then s = Left(s, 31)
    ModelSheetName = s
End Function

' 讀 Portfolio A 欄 ticker 為陣列 (回傳實際筆數 cnt)
Function ReadPortfolioTickers(ByRef tickers() As String) As Integer
    ReadPortfolioTickers = 0
    If Not SheetExists("Portfolio") Then Exit Function

    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets("Portfolio")
    Dim lastRow As Long
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim n As Integer
    n = lastRow - 1
    ReDim tickers(1 To n)
    Dim i As Integer, cnt As Integer, s As String
    cnt = 0
    For i = 1 To n
        s = Trim(CStr(wsP.Cells(i + 1, 1).Value))
        If Len(s) > 0 And UCase(s) <> "TICKER" Then
            cnt = cnt + 1
            tickers(cnt) = s
        End If
    Next i
    If cnt > 0 Then ReDim Preserve tickers(1 To cnt)
    ReadPortfolioTickers = cnt
End Function

' 把指定 sheet 內所有公式凍結為值 (斷開跨 sheet 依賴，避免後續被覆寫資料污染)
Sub FreezeFormulasToValues(sheetName As String)
    If Not SheetExists(sheetName) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ' 先強制計算最新值再凍結
    Application.Calculate
    With ws.UsedRange
        .Value = .Value
    End With
End Sub

' ============================================================
' 為 Portfolio 每檔 ticker 建立獨立 Model 分頁 (M_<ticker>)
' ============================================================
Sub BuildAllModels()
    On Error GoTo errHandler

    Dim tickers() As String, n As Integer
    n = ReadPortfolioTickers(tickers)
    If n = 0 Then
        MsgBox "Portfolio A 欄沒有有效 ticker，請先填入後再執行 (或先跑 BuildPortfolio 建立範本)", vbExclamation
        Exit Sub
    End If

    ' 記下目前主標的以便最後還原
    Dim originalTicker As String
    On Error Resume Next
    originalTicker = ThisWorkbook.Sheets("Dashboard").Range("A2").Value
    On Error GoTo errHandler
    If Len(originalTicker) = 0 Then originalTicker = TICKER_DEFAULT

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 大盤指數抓一次 (Model 公式會 VLOOKUP IndexTWII/IndexOTC)
    Call FetchTaiwanIndices

    Dim failed As String, ok As Integer
    failed = "": ok = 0

    Dim i As Integer
    For i = 1 To n
        Application.StatusBar = "建立 Model " & i & "/" & n & ": " & tickers(i)
        On Error Resume Next
        Call FetchData(tickers(i), "Data")
        If Err.Number = 0 Then
            Dim sheetNm As String
            sheetNm = ModelSheetName(tickers(i))
            Call BuildModel(tickers(i), sheetNm)
            If Err.Number = 0 Then
                ' 關鍵: 立刻凍結公式為值，斷開對 Data 分頁的引用
                ' 否則下一次 FetchData 會覆寫 Data，所有 M_* 都會指向最新 ticker 的資料
                Call FreezeFormulasToValues(sheetNm)
                ok = ok + 1
            Else
                failed = failed & tickers(i) & " "
            End If
        Else
            failed = failed & tickers(i) & " "
        End If
        Err.Clear
        On Error GoTo errHandler
    Next i
    Application.StatusBar = False

    ' 還原 Dashboard 主標的
    Call FetchData(originalTicker, "Data")
    Call BuildModel(originalTicker)          ' 預設寫回 "Model" 分頁
    Call BuildDashboard(originalTicker)

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "完成! 已建立 " & ok & "/" & n & " 個 Model 分頁 (命名: M_<ticker>)"
    If Len(failed) > 0 Then msg = msg & vbCrLf & vbCrLf & "失敗: " & Trim(failed)
    MsgBox msg, vbInformation
    Exit Sub

errHandler:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "BuildAllModels 錯誤: " & Err.Description, vbCritical
End Sub

' ============================================================
' 一鍵清除所有 M_<ticker> 分頁
' ============================================================
Sub DeleteAllModelSheets()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("確定要刪除所有 M_<ticker> 分頁？" & vbCrLf & _
                  "(主要 Model / Dashboard / Data / IndexTWII / IndexOTC / Portfolio 不會被刪)", _
                  vbYesNo + vbQuestion, "確認刪除")
    If resp <> vbYes Then Exit Sub

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim ws As Worksheet, cnt As Integer
    cnt = 0
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 2) = "M_" Then
            ws.Delete
            cnt = cnt + 1
        End If
    Next ws

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "已刪除 " & cnt & " 個 M_<ticker> 分頁", vbInformation
End Sub

Sub BuildPortfolio()
    On Error GoTo errHandler

    ' 1. 確保 Portfolio 分頁存在，第一次跑寫範本
    Dim wsP As Worksheet
    If Not SheetExists("Portfolio") Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "Portfolio"
    End If
    Set wsP = ThisWorkbook.Sheets("Portfolio")

    Dim lastRow As Long
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Or wsP.Cells(1, 1).Value = "" Then
        wsP.Cells(1, 1).Value = "Ticker"
        wsP.Cells(1, 1).Font.Bold = True
        Dim sample As Variant
        sample = Array("2330.TW", "2454.TW", "2317.TW", "2308.TW", "2382.TW", _
                       "3008.TW", "2412.TW", "1301.TW", "1303.TW", "2303.TW", _
                       "2891.TW", "2882.TW", "0050.TW", "0056.TW", "00878.TW")
        Dim k As Integer
        For k = 0 To UBound(sample)
            wsP.Cells(k + 2, 1).Value = sample(k)
        Next k
        Call ApplyDefaultFont(wsP)
        wsP.Columns("A:A").AutoFit
        MsgBox "Portfolio 分頁已建立 15 檔範本。修改 A 欄後再執行 BuildPortfolio。", vbInformation
        Exit Sub
    End If

    Dim n As Integer
    n = lastRow - 1
    Dim tickers() As String
    ReDim tickers(1 To n)
    Dim i As Integer, cnt As Integer
    cnt = 0
    For i = 1 To n
        Dim s As String
        s = Trim(CStr(wsP.Cells(i + 1, 1).Value))
        If Len(s) > 0 Then
            cnt = cnt + 1
            tickers(cnt) = s
        End If
    Next i
    If cnt = 0 Then
        MsgBox "Portfolio A 欄沒有有效 ticker", vbExclamation
        Exit Sub
    End If
    n = cnt
    ReDim Preserve tickers(1 To n)

    ' 2. 記錄目前主標的以便最後還原
    Dim originalTicker As String
    On Error Resume Next
    originalTicker = ThisWorkbook.Sheets("Dashboard").Range("A2").Value
    On Error GoTo errHandler
    If Len(originalTicker) = 0 Then originalTicker = TICKER_DEFAULT

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 3. 抓大盤指數一次 (BuildModel 會用到，避免每檔重抓)
    Call FetchTaiwanIndices

    ' 4. 對每個 ticker 跑模型、同時抽 3 種資料
    Dim dictModel As Object, dictRaw As Object, dictPos As Object
    Set dictModel = CreateObject("Scripting.Dictionary")
    Set dictRaw = CreateObject("Scripting.Dictionary")
    Set dictPos = CreateObject("Scripting.Dictionary")

    For i = 1 To n
        Application.StatusBar = "處理 " & i & "/" & n & ": " & tickers(i)
        Call FetchData(tickers(i), "Data")
        Call BuildModel(tickers(i))
        Call CaptureReturns(i, n, dictModel, dictRaw, dictPos)
    Next i
    Application.StatusBar = False

    ' 5. 寫到 3 個分頁
    Call WriteMatrix("PortfolioReturns", tickers, n, dictModel, "0.0000%")
    Call WriteMatrix("PortfolioRawReturns", tickers, n, dictRaw, "0.0000%")
    Call WriteMatrix("PortfolioPositions", tickers, n, dictPos, "0.0")

    ' 6. 還原主標的
    Call FetchData(originalTicker, "Data")
    Call BuildModel(originalTicker)
    Call BuildDashboard(originalTicker)

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Portfolio 建構完成 (" & n & " 檔) — 3 個矩陣已產生" & vbCrLf & _
           "下一步: 呼叫 ExportAllCSV 一鍵匯出 3 個 CSV", vbInformation
    Exit Sub
errHandler:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "BuildPortfolio 錯誤: " & Err.Description, vbCritical
End Sub

' 從 Model 分頁同步抽 3 種資料 → 3 個 dict
'   dictModel: raw × prevPosition (策略實際曝險報酬)
'   dictRaw:   raw daily return    (純股票報酬，給 Σ 估計用)
'   dictPos:   Position_%          (給 Python 端做 SMA 平滑)
Sub CaptureReturns(idx As Integer, total As Integer, _
                   dictModel As Object, dictRaw As Object, dictPos As Object)
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets("Model")
    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row

    Dim prevClose As Double, prevPos As Double
    prevClose = 0: prevPos = 0

    Dim i As Long
    For i = 2 To lastRow
        Dim d As Variant, c As Variant, posv As Variant
        d = wsM.Cells(i, 1).Value
        c = wsM.Cells(i, 2).Value
        posv = wsM.Cells(i, 20).Value     ' Position_% in T column

        If IsDate(d) And IsNumeric(c) Then
            Dim curClose As Double
            curClose = CDbl(c)
            Dim rawR As Double, modelR As Double
            If prevClose > 0 Then
                rawR = (curClose - prevClose) / prevClose
                modelR = rawR * (prevPos / 100#)
            Else
                rawR = 0: modelR = 0
            End If
            Dim curPos As Double
            If IsNumeric(posv) Then curPos = CDbl(posv) Else curPos = 0

            Dim dKey As Long
            dKey = CLng(CDate(d))

            Call PutToDict(dictModel, dKey, idx, total, modelR)
            Call PutToDict(dictRaw, dKey, idx, total, rawR)
            Call PutToDict(dictPos, dKey, idx, total, curPos)

            prevClose = curClose
            prevPos = curPos
        End If
    Next i
End Sub

Sub PutToDict(d As Object, dKey As Long, idx As Integer, total As Integer, val As Double)
    Dim arr() As Variant
    If d.Exists(dKey) Then
        arr = d(dKey)
    Else
        ReDim arr(1 To total)
        Dim j As Integer
        For j = 1 To total
            arr(j) = ""
        Next j
    End If
    arr(idx) = val
    d(dKey) = arr
End Sub

' 通用矩陣寫入: 把 dict (date->array) 寫到指定 sheet，用指定 numberFormat
Sub WriteMatrix(sheetName As String, tickers() As String, n As Integer, _
                dict As Object, fmt As String)
    If Not SheetExists(sheetName) Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = sheetName
    End If
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "Date"
    Dim i As Integer
    For i = 1 To n
        ws.Cells(1, i + 1).Value = tickers(i)
    Next i
    ws.Rows(1).Font.Bold = True

    Dim keys As Variant
    keys = dict.keys
    Call SortLongArray(keys)

    Dim outArr() As Variant
    ReDim outArr(1 To UBound(keys) + 1, 1 To n + 1)

    Dim r As Long
    For r = 0 To UBound(keys)
        outArr(r + 1, 1) = CDate(keys(r))
        Dim arrv As Variant
        arrv = dict(keys(r))
        Dim j As Integer
        For j = 1 To n
            outArr(r + 1, j + 1) = arrv(j)
        Next j
    Next r

    ws.Range(ws.Cells(2, 1), ws.Cells(UBound(keys) + 2, n + 1)).Value = outArr
    ws.Columns("A:A").NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(2, 2), ws.Cells(UBound(keys) + 2, n + 1)).NumberFormat = fmt
    Call ApplyDefaultFont(ws)
    ws.Columns("A:Z").AutoFit
End Sub

Sub SortLongArray(arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    Dim hi As Long
    hi = UBound(arr)
    For i = 0 To hi - 1
        For j = i + 1 To hi
            If arr(i) > arr(j) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

' 單一 sheet 匯出 (相容舊呼叫)
Sub ExportPortfolioCSV()
    Call ExportSheetToCSV("PortfolioReturns", "portfolio_returns.csv")
End Sub

' 一鍵匯出 3 個 CSV
Sub ExportAllCSV()
    Call ExportSheetToCSV("PortfolioReturns", "portfolio_returns.csv")
    Call ExportSheetToCSV("PortfolioRawReturns", "portfolio_raw_returns.csv")
    Call ExportSheetToCSV("PortfolioPositions", "portfolio_positions.csv")
    MsgBox "3 個 CSV 已匯出到: " & GetExportDir() & vbCrLf & vbCrLf & _
           "下一步: python portfolio_optimizer.py", vbInformation
End Sub

' 通用單 sheet 匯出
Sub ExportSheetToCSV(sheetName As String, csvName As String)
    If Not SheetExists(sheetName) Then
        MsgBox "找不到 " & sheetName & "，請先執行 BuildPortfolio", vbExclamation
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Dim csvPath As String
    csvPath = GetExportDir() & "\" & csvName

    On Error GoTo errHandler
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim wsCopy As Worksheet
    ws.Copy
    Set wsCopy = ActiveWorkbook.Sheets(1)
    wsCopy.Columns("A:A").NumberFormat = "yyyy-mm-dd"
    Dim usedCol As Long
    usedCol = wsCopy.Cells(1, wsCopy.Columns.Count).End(xlToLeft).Column
    If usedCol >= 2 Then
        wsCopy.Range(wsCopy.Cells(2, 2), _
            wsCopy.Cells(wsCopy.UsedRange.Rows.Count, usedCol)).NumberFormat = "0.00000000"
    End If

    ActiveWorkbook.SaveAs Filename:=csvPath, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
errHandler:
    On Error Resume Next
    ActiveWorkbook.Close SaveChanges:=False
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "匯出 " & csvName & " 失敗 #" & Err.Number & ": " & Err.Description, vbCritical
End Sub

Function GetExportDir() As String
    If Len(ThisWorkbook.path) > 0 Then
        GetExportDir = ThisWorkbook.path
    Else
        GetExportDir = Environ("TEMP")
    End If
End Function
