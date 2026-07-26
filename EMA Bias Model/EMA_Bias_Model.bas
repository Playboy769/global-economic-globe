Attribute VB_Name = "EMA_Bias_Model"
Option Explicit

' ============================================================
'  EMA Bias 加倉模型 (14-45 天波段)
'  指標: EMA6 / EMA13 / EMA30 / SMA60 + 量能 + 大盤/櫃買拖累
'  EMA 65% (排列25 + 乖離40) + 量能 35% - 指數拖累 (最多 -25)
'  Mode: 趨勢加碼 / 淺回檔 / 深回檔 / 趨勢轉弱 / 整理破壞
'  資料來源: Yahoo Finance v8 chart API (yfinance 底層相同端點)
' ============================================================

Public Const TICKER_DEFAULT As String = "2330.TW"
' Yahoo range=1y 約 252 個交易日 (約 1 年)，足夠 SMA60 + 緩衝
Public Const LOOKBACK_YEARS As Integer = 1
Public Const IDX_TWII As String = "^TWII"
Public Const IDX_OTC As String = "^TWOII"

' 統一字體設定
Public Const DEFAULT_FONT_NAME As String = "Yu Gothic"
Public Const DEFAULT_FONT_SIZE As Integer = 12

' 套用預設字體到整個工作表 (保留粗體/顏色等屬性，只統一字型與字級)
Public Sub ApplyDefaultFont(ws As Worksheet)
    With ws.Cells.Font
        .Name = DEFAULT_FONT_NAME
        .Size = DEFAULT_FONT_SIZE
    End With
End Sub

' 內部 helper: 對指定 sheet 套用「全表置中 + 首列凍結窗格」
' FreezePanes 是 window 層級設定，需要 Activate 該 sheet 才能套用
Public Sub ApplyModelSheetFormatting(ws As Worksheet)
    ' 1. 全表水平與垂直置中
    With ws.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 2. 凍結首列 (Row 1)
    Dim prevSheet As Object
    Set prevSheet = ActiveSheet
    ws.Activate
    With ActiveWindow
        .FreezePanes = False     ' 先解除舊的設定避免衝突
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    ' 還原原本啟用中的 sheet (不打擾使用者)
    On Error Resume Next
    prevSheet.Activate
    On Error GoTo 0
End Sub

' ============================================================
' 【獨立巨集】套用版面 (置中 + 凍結首列) 到「目前啟用」的 sheet
' 用途: 在某個分頁上手動微調後，只想單獨對這個分頁套用版面
' ============================================================
Sub FormatCurrentSheet()
    If ActiveSheet Is Nothing Then
        MsgBox "沒有任何啟用中的 sheet", vbExclamation
        Exit Sub
    End If
    Call ApplyModelSheetFormatting(ActiveSheet)
    MsgBox "已套用版面到: " & ActiveSheet.Name & vbCrLf & _
           "(全表置中 + 首列凍結)", vbInformation
End Sub

' ============================================================
' 【獨立巨集】一次套用版面到所有 Model 分頁
' 對象: "Model" + 所有 "M_*" 分頁 (BuildAllModels 產出的)
' ============================================================
Sub FormatAllModelSheets()
    Application.ScreenUpdating = False

    Dim ws As Worksheet, cnt As Integer
    cnt = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Model" Or Left(ws.Name, 2) = "M_" Then
            Call ApplyModelSheetFormatting(ws)
            cnt = cnt + 1
        End If
    Next ws

    Application.ScreenUpdating = True
    MsgBox "已套用版面到 " & cnt & " 個 Model 分頁" & vbCrLf & _
           "(全表置中 + 首列凍結)", vbInformation
End Sub

' ---------- Entry point ----------
Sub Setup()
    On Error GoTo errHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Call BuildSheets
    Call FetchData(TICKER_DEFAULT, "Data")
    Call FetchTaiwanIndices
    Call BuildModel(TICKER_DEFAULT)
    Call BuildDashboard(TICKER_DEFAULT)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets("Dashboard").Activate
    MsgBox "完成！在 Dashboard A2 修改代碼即可換股。", vbInformation
    Exit Sub
errHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Setup 錯誤: " & Err.Description, vbCritical
End Sub

' 換標的時呼叫: 例如 Call Refresh("2454.TW") 看聯發科
Sub Refresh(Optional ticker As String = "")
    On Error GoTo errHandler
    If ticker = "" Then ticker = TICKER_DEFAULT
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Call BuildSheets
    Call FetchData(ticker, "Data")
    Call FetchTaiwanIndices
    Call BuildModel(ticker)
    Call BuildDashboard(ticker)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
errHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Refresh 錯誤 (" & ticker & "): " & Err.Description, vbCritical
End Sub

' 由按鈕或 Worksheet_Change 呼叫: 讀 Dashboard!A2 後重算
Sub RunFromInput()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Dim t As String
    t = Trim(CStr(ws.Range("A2").Value))
    If Len(t) = 0 Then
        MsgBox "請在 A2 輸入代碼 (例如 2330.TW)", vbExclamation
        Exit Sub
    End If
    Call Refresh(t)
End Sub

' ---------- Sheet skeleton ----------
Sub BuildSheets()
    Dim sheetNames As Variant
    sheetNames = Array("Data", "IndexTWII", "IndexOTC", "Model", "Dashboard")
    Dim i As Integer
    For i = 0 To UBound(sheetNames)
        If Not SheetExists(CStr(sheetNames(i))) Then
            ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = CStr(sheetNames(i))
        End If
    Next i
End Sub

Function SheetExists(name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' ---------- Data fetch (Yahoo Finance v8 chart JSON) ----------
Sub FetchData(ticker As String, sheetName As String)
    Dim url As String
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & ticker & _
          "?range=" & LOOKBACK_YEARS & "y&interval=1d"

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    ' Yahoo 會擋掉沒有 UA 的請求
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    http.send

    If http.Status <> 200 Then
        MsgBox "Yahoo Finance 抓取失敗 (" & ticker & " HTTP " & http.Status & ")" & vbCrLf & _
               "檢查 ticker 格式: 台股需加 .TW (例如 2330.TW)", vbCritical
        Exit Sub
    End If

    Dim json As String
    json = http.responseText

    Dim ts As Variant, op As Variant, hi As Variant, lo As Variant, cl As Variant, vol As Variant
    ts = ExtractJsonArray(json, """timestamp"":[")
    op = ExtractJsonArray(json, """open"":[")
    hi = ExtractJsonArray(json, """high"":[")
    lo = ExtractJsonArray(json, """low"":[")
    cl = ExtractJsonArray(json, """close"":[")
    vol = ExtractJsonArray(json, """volume"":[")

    If IsEmpty(ts) Or IsEmpty(cl) Then
        MsgBox "JSON 解析失敗 (" & ticker & ")，檢查 ticker", vbCritical
        Exit Sub
    End If

    Dim n As Long
    n = UBound(ts)

    ' 保險: 分頁不存在就先建
    If Not SheetExists(sheetName) Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = sheetName
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Cells.Clear

    Dim arr() As Variant
    ReDim arr(1 To n + 2, 1 To 6)
    arr(1, 1) = "Date"
    arr(1, 2) = "Open"
    arr(1, 3) = "High"
    arr(1, 4) = "Low"
    arr(1, 5) = "Close"
    arr(1, 6) = "Volume"

    Dim epoch As Date
    epoch = DateSerial(1970, 1, 1)

    Dim i As Long, outRow As Long
    outRow = 1
    For i = 0 To n
        ' Yahoo 對非交易日會回傳 null，過濾掉
        If Not IsEmpty(cl(i)) And cl(i) > 0 Then
            outRow = outRow + 1
            arr(outRow, 1) = DateAdd("s", CDbl(ts(i)), epoch)
            arr(outRow, 2) = SafeNum(op(i))
            arr(outRow, 3) = SafeNum(hi(i))
            arr(outRow, 4) = SafeNum(lo(i))
            arr(outRow, 5) = SafeNum(cl(i))
            arr(outRow, 6) = SafeNum(vol(i))
        End If
    Next i

    ws.Range(ws.Cells(1, 1), ws.Cells(outRow, 6)).Value = arr
    ws.Columns("A:A").NumberFormat = "yyyy-mm-dd"
    Call ApplyDefaultFont(ws)
    ws.Columns("A:F").AutoFit
End Sub

' 從 JSON 字串擷取 marker 之後到 ] 的數字陣列 (含 null)
Function ExtractJsonArray(json As String, marker As String) As Variant
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, json, marker)
    If startPos = 0 Then
        ExtractJsonArray = Empty
        Exit Function
    End If
    startPos = startPos + Len(marker)
    endPos = InStr(startPos, json, "]")
    If endPos = 0 Then
        ExtractJsonArray = Empty
        Exit Function
    End If

    Dim body As String
    body = Mid$(json, startPos, endPos - startPos)
    Dim parts() As String
    parts = Split(body, ",")

    Dim arr() As Variant
    ReDim arr(0 To UBound(parts))
    Dim i As Long
    For i = 0 To UBound(parts)
        If parts(i) = "null" Or Len(Trim(parts(i))) = 0 Then
            arr(i) = Empty
        Else
            arr(i) = CDbl(parts(i))
        End If
    Next i
    ExtractJsonArray = arr
End Function

Function SafeNum(v As Variant) As Double
    If IsEmpty(v) Then SafeNum = 0 Else SafeNum = CDbl(v)
End Function

' ---------- Taiwan index fetch & decline-day counter ----------
Sub FetchTaiwanIndices()
    Call FetchData(IDX_TWII, "IndexTWII")
    Call FetchData(IDX_OTC, "IndexOTC")
    Call BuildIndexSheet("IndexTWII")
    Call BuildIndexSheet("IndexOTC")
End Sub

' 在指數 Data 後附加 EMA20、DeclineFlag、DeclineDays 三欄
' DeclineFlag: 1 若 EMA20 比前一日低 (短期向下傾斜)
' DeclineDays: 連續下降天數，回升即歸零
Sub BuildIndexSheet(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 25 Then Exit Sub

    ws.Cells(1, 7).Value = "EMA20"
    ws.Cells(1, 8).Value = "DeclineFlag"
    ws.Cells(1, 9).Value = "DeclineDays"
    ws.Rows(1).Font.Bold = True

    Dim i As Long
    For i = 2 To lastRow
        If i = 2 Then
            ws.Cells(i, 7).Formula = "=E" & i
            ws.Cells(i, 8).Value = 0
            ws.Cells(i, 9).Value = 0
        Else
            ws.Cells(i, 7).Formula = "=(2/21)*E" & i & "+(19/21)*G" & i - 1
            ws.Cells(i, 8).Formula = "=IF(G" & i & "<G" & i - 1 & ",1,0)"
            ws.Cells(i, 9).Formula = "=IF(H" & i & "=1,I" & i - 1 & "+1,0)"
        End If
    Next i

    ws.Columns("G:G").NumberFormat = "#,##0.00"
    Call ApplyDefaultFont(ws)
    ws.Columns("A:I").AutoFit
End Sub

Function IsTaiwanTicker(t As String) As Boolean
    Dim u As String
    u = UCase(t)
    IsTaiwanTicker = (Right(u, 3) = ".TW") Or (Right(u, 4) = ".TWO")
End Function

' ---------- Model sheet (indicators + scores via formulas) ----------
' modelSheetName 可選: 預設 "Model"。傳入 "M_2330.TW" 等可建立每檔獨立分頁
Sub BuildModel(ticker As String, Optional modelSheetName As String = "Model")
    Dim wsD As Worksheet, wsM As Worksheet
    Set wsD = ThisWorkbook.Sheets("Data")
    If Not SheetExists(modelSheetName) Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = modelSheetName
    End If
    Set wsM = ThisWorkbook.Sheets(modelSheetName)

    Dim lastRow As Long
    lastRow = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    If lastRow < 70 Then
        MsgBox "資料筆數不足 (" & lastRow & ")，至少需 70 個交易日。", vbCritical
        Exit Sub
    End If

    Dim isTW As Boolean
    isTW = IsTaiwanTicker(ticker)

    wsM.Cells.Clear

    Dim headers As Variant
    headers = Array("Date", "Close", "Volume", _
                    "EMA6", "EMA13", "EMA30", "SMA60", "VolSMA20", _
                    "Bias_SMA60_%", "OrderScore", "DeviationScore", _
                    "VolRatio", "VolScore", _
                    "TWII_DD", "OTC_DD", "IndexPenalty", _
                    "Ret5D_%", "PanicFlag", _
                    "TotalScore", _
                    "High30", "Pullback_%", "PullbackScore")
    Dim k As Integer
    For k = 0 To UBound(headers)
        wsM.Cells(1, k + 1).Value = headers(k)
    Next k
    wsM.Rows(1).Font.Bold = True

    Dim i As Long
    For i = 2 To lastRow
        wsM.Cells(i, 1).Formula = "=Data!A" & i
        wsM.Cells(i, 2).Formula = "=Data!E" & i
        wsM.Cells(i, 3).Formula = "=Data!F" & i

        If i = 2 Then
            wsM.Cells(i, 4).Formula = "=B" & i
            wsM.Cells(i, 5).Formula = "=B" & i
            wsM.Cells(i, 6).Formula = "=B" & i
        Else
            wsM.Cells(i, 4).Formula = "=(2/7)*B" & i & "+(5/7)*D" & i - 1
            wsM.Cells(i, 5).Formula = "=(2/14)*B" & i & "+(12/14)*E" & i - 1
            wsM.Cells(i, 6).Formula = "=(2/31)*B" & i & "+(29/31)*F" & i - 1
        End If

        If i >= 61 Then
            wsM.Cells(i, 7).Formula = "=AVERAGE(B" & i - 59 & ":B" & i & ")"
        End If
        If i >= 21 Then
            wsM.Cells(i, 8).Formula = "=AVERAGE(C" & i - 19 & ":C" & i & ")"
        End If

        wsM.Cells(i, 9).Formula = _
            "=IF(G" & i & "="""","""",(B" & i & "-G" & i & ")/G" & i & "*100)"

        wsM.Cells(i, 10).Formula = _
            "=IF(G" & i & "="""","""",IF(B" & i & ">D" & i & ",6,0)+IF(D" & i & ">E" & i & ",6,0)+IF(E" & i & ">F" & i & ",6,0)+IF(F" & i & ">G" & i & ",7,0))"

        ' DeviationScore: SMA60 乖離分 (0~20)，越遠越好
        wsM.Cells(i, 11).Formula = _
            "=IF(I" & i & "="""",""""," & _
            "IF(I" & i & ">20,20,IF(I" & i & ">15,17,IF(I" & i & ">10,14,IF(I" & i & ">5,11,IF(I" & i & ">0,7,3))))))"

        wsM.Cells(i, 12).Formula = _
            "=IF(H" & i & "="""","""",C" & i & "/H" & i & ")"

        wsM.Cells(i, 13).Formula = _
            "=IF(L" & i & "="""","""",IF(L" & i & ">4,10,IF(L" & i & ">2.5,22,IF(L" & i & ">1.5,35,IF(L" & i & ">1,26,IF(L" & i & ">0.7,15,5))))))"

        ' 指數連續下降天數 (僅對台股 ticker 套用)
        If isTW Then
            wsM.Cells(i, 14).Formula = "=IFERROR(VLOOKUP(A" & i & ",IndexTWII!A:I,9,FALSE),0)"
            wsM.Cells(i, 15).Formula = "=IFERROR(VLOOKUP(A" & i & ",IndexOTC!A:I,9,FALSE),0)"
            ' IndexPenalty 無 cap，累計可超過 60 → 觸發恐慌條款
            wsM.Cells(i, 16).Formula = "=N" & i & "*0.75+O" & i & "*0.5"
        Else
            wsM.Cells(i, 14).Value = 0
            wsM.Cells(i, 15).Value = 0
            wsM.Cells(i, 16).Value = 0
        End If

        ' 近 5 日漲跌幅 (%)
        If i >= 7 Then
            wsM.Cells(i, 17).Formula = "=(B" & i & "-B" & i - 5 & ")/B" & i - 5 & "*100"
        End If

        ' PanicFlag: IndexPenalty > 60 且 5 日跌幅 < -15%
        wsM.Cells(i, 18).Formula = _
            "=IF(Q" & i & "="""",0,IF(AND(P" & i & ">60,Q" & i & "<-15),1,0))"

        ' TotalScore: 排列25 + 乖離20 + 回撤20 + 量能35 = 100
        ' PanicFlag=1 時 VolScore 雙倍 + 跳過 IndexPenalty；否則含拖累扣分
        ' PullbackScore 現在位於 V 欄 (原 Y 欄已左移)
        wsM.Cells(i, 19).Formula = _
            "=IF(OR(J" & i & "="""",K" & i & "="""",M" & i & "="""",V" & i & "=""""),""""," & _
            "MAX(0,IF(R" & i & "=1," & _
                "J" & i & "+K" & i & "+V" & i & "+M" & i & "*2," & _
                "J" & i & "+K" & i & "+V" & i & "+M" & i & "-P" & i & ")))"

        ' T (col 20) High30: 近 30 交易日最高價 (取 Data!C 高價欄)，需 i>=31 才滿 30 筆
        If i >= 31 Then
            wsM.Cells(i, 20).Formula = "=MAX(Data!C" & i - 29 & ":Data!C" & i & ")"
        End If

        ' U (col 21) Pullback_%: 從 30 日高點回撤 (= (Close - High30) / High30 * 100，<=0)
        wsM.Cells(i, 21).Formula = _
            "=IF(T" & i & "="""","""",(B" & i & "-T" & i & ")/T" & i & "*100)"

        ' V (col 22) PullbackScore (0~20): 30 日高點回撤越小分越高
        wsM.Cells(i, 22).Formula = _
            "=IF(U" & i & "="""",""""," & _
            "IF(-U" & i & "<=2,20,IF(-U" & i & "<=5,16,IF(-U" & i & "<=10,12,IF(-U" & i & "<=15,8,IF(-U" & i & "<=25,4,1))))))"
    Next i

    wsM.Columns("A:A").NumberFormat = "yyyy-mm-dd"
    wsM.Columns("B:H").NumberFormat = "#,##0.00"
    wsM.Columns("I:I").NumberFormat = "0.00"
    wsM.Columns("J:K").NumberFormat = "0.0"
    wsM.Columns("L:L").NumberFormat = "0.00"
    wsM.Columns("M:M").NumberFormat = "0.0"
    wsM.Columns("N:O").NumberFormat = "0"
    wsM.Columns("P:P").NumberFormat = "0.0"
    wsM.Columns("Q:Q").NumberFormat = "0.00"
    wsM.Columns("R:R").NumberFormat = "0"
    wsM.Columns("S:S").NumberFormat = "0.0"
    wsM.Columns("T:T").NumberFormat = "#,##0.00"       ' High30
    wsM.Columns("U:U").NumberFormat = "0.00;[Red]-0.00" ' Pullback_%
    wsM.Columns("V:V").NumberFormat = "0"              ' PullbackScore
    Call ApplyDefaultFont(wsM)
    wsM.Columns("A:V").AutoFit

    ' TotalScore 顏色帶 (移到 S 欄；原本套在 Position_% T 欄但已移除)
    With wsM.Range("S2:S" & lastRow)
        .FormatConditions.Delete
        .FormatConditions.AddColorScale ColorScaleType:=3
        .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
        .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .FormatConditions(1).ColorScaleCriteria(2).Value = 50
        .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
    End With

    ' IndexPenalty 顏色 (越紅越糟)，0~75 區間
    With wsM.Range("P2:P" & lastRow)
        .FormatConditions.Delete
        .FormatConditions.AddColorScale ColorScaleType:=2
        .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueNumber
        .FormatConditions(1).ColorScaleCriteria(1).Value = 0
        .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(255, 255, 255)
        .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueNumber
        .FormatConditions(1).ColorScaleCriteria(2).Value = 75
        .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(192, 0, 0)
    End With

    ' PanicFlag 高亮 (=1 → 整列填黃)
    With wsM.Range("R2:R" & lastRow)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="1"
        .FormatConditions(1).Interior.Color = RGB(255, 217, 102)
        .FormatConditions(1).Font.Bold = True
        .FormatConditions(1).Font.Color = RGB(192, 0, 0)
    End With

    ' 版面 (置中 + 首列凍結) 已拆成獨立巨集，需要時再執行:
    '   FormatAllModelSheets  → 一次套用到所有 Model 與 M_<ticker>
    '   FormatCurrentSheet    → 只套用到當前啟用的 sheet
End Sub

' ---------- Dashboard ----------
Sub BuildDashboard(ticker As String)
    Dim ws As Worksheet, wsM As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Set wsM = ThisWorkbook.Sheets("Model")
    ws.Cells.Clear

    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row

    With ws.Range("A1")
        .Value = ticker
        .Font.Size = DEFAULT_FONT_SIZE
        .Font.Bold = True
    End With

    ' A2 = 輸入框 (黃底紅框)，使用者改值後按 Enter 即透過 Worksheet_Change 自動重算
    With ws.Range("A2")
        .Value = ticker
        .Interior.Color = RGB(255, 255, 153)
        .Font.Bold = True
        .Font.Size = DEFAULT_FONT_SIZE
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(192, 0, 0)
        .Borders.Weight = xlMedium
    End With
    With ws.Range("B2")
        .Value = "← 改 A2 代碼按 Enter 即自動重算 (需安裝 Worksheet_Change 事件)"
        .Font.Italic = True
        .Font.Color = RGB(100, 100, 100)
    End With

    ' 清除所有 Shape (包含舊版可能殘留的「套用 A2 代碼」按鈕)
    Dim sh As Shape
    For Each sh In ws.Shapes
        sh.Delete
    Next sh

    Dim labels As Variant, refs As Variant
    labels = Array("最新日期", "收盤價", _
                   "EMA6", "EMA13", "EMA30", "SMA60", _
                   "SMA60 乖離率 (%)", "30 日高點", "30 日高點回撤 (%)", _
                   "排列分 (/25)", "乖離分 (/20)", "回撤分 (/20)", _
                   "量能比 (vs 20MA)", "量能分 (/35)", _
                   "加權連跌天數", "櫃買連跌天數", "指數拖累 (扣分)", _
                   "近 5 日漲跌 (%)", "恐慌旗標", _
                   "總分")
    ' Model 欄位左移後: High30=T, Pullback=U, PullbackScore=V
    refs = Array("A", "B", "D", "E", "F", "G", "I", "T", "U", _
                 "J", "K", "V", "L", "M", _
                 "N", "O", "P", "Q", "R", _
                 "S")

    Dim r As Integer
    For r = 0 To UBound(labels)
        ws.Cells(r + 3, 1).Value = labels(r)
        ws.Cells(r + 3, 2).Formula = "=Model!" & refs(r) & lastRow
    Next r

    ' 列號 (共 20 項，列 3~22):
    '   3 最新日期  4 收盤價  5-8 EMA/SMA  9 乖離率
    '  10 高點  11 回撤  12 排列分  13 乖離分  14 回撤分  15 量能比  16 量能分
    '  17 TWII連跌  18 OTC連跌  19 指數拖累  20 近5日  21 恐慌
    '  22 總分
    ws.Range("A3:A22").Font.Bold = True
    ws.Range("B3").NumberFormat = "yyyy-mm-dd"
    ws.Range("B4:B8").NumberFormat = "#,##0.00"
    ws.Range("B9").NumberFormat = "0.00"
    ws.Range("B10").NumberFormat = "#,##0.00"         ' 30 日高點
    ws.Range("B11").NumberFormat = "0.00;[Red]-0.00"  ' 30 日高點回撤 (%)
    ws.Range("B12:B14").NumberFormat = "0"            ' 排列/乖離/回撤分
    ws.Range("B15").NumberFormat = "0.00"             ' 量能比
    ws.Range("B16").NumberFormat = "0"                ' 量能分
    ws.Range("B17:B18").NumberFormat = "0"            ' 連跌天數
    ws.Range("B19").NumberFormat = "0.0;[Red]-0.0"    ' 指數拖累
    ws.Range("B20").NumberFormat = "0.00"             ' 近5日
    ws.Range("B21").NumberFormat = "0"                ' 恐慌旗標
    ws.Range("B22").NumberFormat = "0.0"              ' 總分

    ' 高亮: 總分 (列 22)
    ws.Range("B22").Font.Bold = True
    ws.Range("B22").Font.Color = RGB(192, 0, 0)
    ws.Range("B22").Font.Size = DEFAULT_FONT_SIZE

    ' PanicFlag = 1 整列警示 (列 21)
    With ws.Range("B21")
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="1"
        .FormatConditions(1).Interior.Color = RGB(255, 217, 102)
        .FormatConditions(1).Font.Bold = True
        .FormatConditions(1).Font.Color = RGB(192, 0, 0)
    End With

    ws.Range("D3").Value = "計分規則"
    ws.Range("D3").Font.Bold = True
    ws.Range("D3").Font.Size = DEFAULT_FONT_SIZE

    ' 規則說明 — 逐行寫入避免 VBA 24 個續行符上限
    Dim rr As Integer
    rr = 4
    ws.Cells(rr, 4) = "[EMA 65%] = 排列分 25 + 乖離分 20 + 回撤分 20": rr = rr + 1
    ws.Cells(rr, 4) = "  排列分: Price>EMA6=6, EMA6>EMA13=6, EMA13>EMA30=6, EMA30>SMA60=7": rr = rr + 1
    ws.Cells(rr, 4) = "  乖離分 (對 SMA60，越遠越好):": rr = rr + 1
    ws.Cells(rr, 4) = "    0~5%=7、5~10%=11、10~15%=14、15~20%=17、>20%=20、<0%=3": rr = rr + 1
    ws.Cells(rr, 4) = "  回撤分 (近 30 交易日高點回撤，越小越好):": rr = rr + 1
    ws.Cells(rr, 4) = "    0~2%=20、2~5%=16、5~10%=12、10~15%=8、15~25%=4、>25%=1": rr = rr + 1
    rr = rr + 1
    ws.Cells(rr, 4) = "[量能 35%] (今日量 / 20MA):": rr = rr + 1
    ws.Cells(rr, 4) = "    1.5~2.5x=35 (理想)、2.5~4=22、1~1.5=26、0.7~1=15、>4=10、<0.7=5": rr = rr + 1
    rr = rr + 1
    ws.Cells(rr, 4) = "[指數拖累] 加權/櫃買 EMA20 連續下降天數 (僅台股 .TW)": rr = rr + 1
    ws.Cells(rr, 4) = "    加權: 連跌天數 × 0.75 (無上限)": rr = rr + 1
    ws.Cells(rr, 4) = "    櫃買: 連跌天數 × 0.5  (無上限)": rr = rr + 1
    ws.Cells(rr, 4) = "    指數連跌愈久、扣愈多，反映系統性風險": rr = rr + 1
    rr = rr + 1
    ws.Cells(rr, 4) = "[恐慌條款 ★] 當 IndexPenalty > 60 且近 5 日跌幅 < -15%:": rr = rr + 1
    ws.Cells(rr, 4) = "    PanicFlag = 1 (整列高亮黃底紅字)": rr = rr + 1
    ws.Cells(rr, 4) = "    重評分: 量能分 × 2 (max 70)，跳過 IndexPenalty 扣分": rr = rr + 1
    ws.Cells(rr, 4) = "    觸發時機: 大型熊市末端急殺 (例如 2020/3 covid 崩盤)": rr = rr + 1
    rr = rr + 1
    ws.Cells(rr, 4) = "[總分] = OrderScore + DeviationScore + PullbackScore + VolScore - IndexPenalty (?0)": rr = rr + 1
    ws.Cells(rr, 4) = "  滿分 100，越高代表加碼條件越強"

    Call ApplyDefaultFont(ws)
    ws.Columns("A:B").AutoFit
    ws.Columns("D:D").ColumnWidth = 65
End Sub
