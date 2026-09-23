Attribute VB_Name = "modMISPrice"
Option Explicit

' ================================================================
'  MIS PRICE FETCHER  v1.0  (2026-09-23)
'
'  Public Sub      PrefetchMISPrices(tickers() As String)
'  Public Function GetMISPrice(ticker As String) As Double
'
'  台灣證交所「盤中即時資訊」端點（mis.twse.com.tw），非官方公開 API、
'  無 SLA、無官方文件。實測確認：
'    - 不需要 session cookie（index.jsp 現在回 404，直接 GET 即可拿到資料）
'    - 一次查詢可用 "|" 串接多檔混合上市(tse_)/上櫃(otc_)，陣列回應順序
'      與輸入順序一致
'    - 市場前綴配錯的代號回傳空殼物件（無 "@" 欄位），只能用「陣列位置」
'      比對，不能用內容反查
'    - 頂層 "z" 欄位只在剛好有新成交的瞬間才有值，兩次查詢之間常變回 "-"；
'      要抓「最新成交價」一律讀巢狀 "trade":{"z":...} 欄位
'  社群經驗與伺服器自報的 "userDelay":5000 一致指向「5 秒內不要超過 3 次
'  請求」。本模組的因應方式：
'    1) 一次 UP 只發 1 次 HTTP 請求，所有代號用 "|" 串進同一個查詢
'       （由 PortfolioDashboard_v3.RebuildPortfolioDashboard 呼叫）
'    2) 加最短查詢間隔 MIN_QUERY_GAP_SEC，避免使用者連續快速按 UP
'       觸發連續請求
'    3) 失敗（HTTP 錯誤、被擋、非交易時段抓不到）時快取直接留空，
'       GetMISPrice 回 0，呼叫端（Attach.GetStockPrice）會自動退 Yahoo，
'       不重試、不硬撐
'
'  美股不經此模組；只處理 .TW / .TWO 代號。
' ================================================================

Private Const MIS_BASE_URL As String = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch="
Private Const MIN_QUERY_GAP_SEC As Double = 5   ' 對齊伺服器 userDelay / 社群「5秒3次」經驗

Private g_misCache     As Object    ' Dictionary: 代號(UCase, 含 .TW/.TWO) -> 成交價
Private g_lastQueryAt  As Double    ' 上次實際發出 HTTP 請求的 Timer() 秒數

' ================================================================
'  PUBLIC: 一次批次查詢，結果存入模組層快取
'  tickers() 需已是 UCase、含 .TW 或 .TWO 後綴（呼叫端負責正規化）
' ================================================================
Public Sub PrefetchMISPrices(tickers() As String)
    Dim n As Long
    On Error Resume Next
    n = UBound(tickers) - LBound(tickers) + 1
    On Error GoTo 0
    If n <= 0 Then Exit Sub

    ' 節流：距離上次真正發出請求不到 MIN_QUERY_GAP_SEC 秒就跳過，
    ' 保留現有快取（可能是空的，也可能是上一輪的結果），不清空、不重抓。
    If g_lastQueryAt > 0 Then
        Dim gap As Double: gap = Timer - g_lastQueryAt
        If gap >= 0 And gap < MIN_QUERY_GAP_SEC Then Exit Sub
    End If

    ' 每次真正查詢前重建乾淨快取——避免這輪查詢失敗時，舊快取的過期價格
    ' 被 GetMISPrice 誤當成「這輪」的有效結果一直沿用下去。
    Set g_misCache = CreateObject("Scripting.Dictionary")

    Dim qParts()   As String: ReDim qParts(0 To n - 1)
    Dim origTicker() As String: ReDim origTicker(0 To n - 1)
    Dim isOtc()    As Boolean: ReDim isOtc(0 To n - 1)

    Dim i As Long, j As Long: j = 0
    For i = LBound(tickers) To UBound(tickers)
        Dim t As String: t = UCase(Trim(tickers(i)))
        If Len(t) = 0 Then GoTo ContinueI
        Dim code As String: code = StripSuffixMIS(t)
        If Len(code) = 0 Then GoTo ContinueI
        Dim otc As Boolean: otc = (InStr(t, ".TWO") > 0)
        origTicker(j) = t
        isOtc(j) = otc
        qParts(j) = IIf(otc, "otc_", "tse_") & code & ".tw"
        j = j + 1
ContinueI:
    Next i
    If j = 0 Then Exit Sub

    ReDim Preserve qParts(0 To j - 1)
    ReDim Preserve origTicker(0 To j - 1)
    ReDim Preserve isOtc(0 To j - 1)

    Dim resp As String: resp = MisHttpGet(MIS_BASE_URL & Join(qParts, "|"))
    g_lastQueryAt = Timer
    If Len(resp) = 0 Then Exit Sub   ' 查詢失敗：快取留空，讓 GetMISPrice 全部回 0 退 Yahoo

    Dim entries As Collection: Set entries = SplitMsgArray(resp)

    Dim k As Long
    For k = 1 To entries.Count
        If k - 1 > UBound(origTicker) Then Exit For
        Dim entryStr As String: entryStr = entries(k)
        Dim px As Double: px = ExtractTradeZ(entryStr)

        If px = 0 Then
            ' 市場前綴可能配錯（極少見，.TW/.TWO 後綴在本活頁簿裡本來就是
            ' 市場真相），換另一邊單檔重試一次
            Dim altPrefix As String: altPrefix = IIf(isOtc(k - 1), "tse_", "otc_")
            Dim altCode As String: altCode = StripSuffixMIS(origTicker(k - 1))
            Dim altResp As String: altResp = MisHttpGet(MIS_BASE_URL & altPrefix & altCode & ".tw")
            If Len(altResp) > 0 Then
                Dim altEntries As Collection: Set altEntries = SplitMsgArray(altResp)
                If altEntries.Count >= 1 Then px = ExtractTradeZ(altEntries(1))
            End If
        End If

        If px > 0 Then g_misCache(origTicker(k - 1)) = px
    Next k
End Sub

' ================================================================
'  PUBLIC: 從這一輪的批次快取讀價，沒有 prefetch 過或抓不到一律回 0
' ================================================================
Public Function GetMISPrice(ticker As String) As Double
    GetMISPrice = 0
    If g_misCache Is Nothing Then Exit Function
    Dim t As String: t = UCase(Trim(ticker))
    If g_misCache.Exists(t) Then GetMISPrice = g_misCache(t)
End Function

' ================================================================
'  JSON 解析（手動字串處理，不依賴外部 JSON 函式庫——與
'  TaiwanPriceFetcher.ParsePrice 同一套做法）
' ================================================================

' 把 "msgArray":[ {...}, {...} ] 拆成個別物件字串，用大括號深度計數，
' 才不會被每個物件內巢狀的 "trade":{...} 切錯。
Private Function SplitMsgArray(json As String) As Collection
    Dim result As New Collection
    Dim marker As String: marker = """msgArray"":["
    Dim startPos As Long: startPos = InStr(json, marker)
    If startPos = 0 Then Set SplitMsgArray = result: Exit Function
    startPos = startPos + Len(marker)

    Dim depth As Long: depth = 0
    Dim entryStart As Long: entryStart = 0
    Dim i As Long
    For i = startPos To Len(json)
        Dim ch As String: ch = Mid(json, i, 1)
        If ch = "{" Then
            If depth = 0 Then entryStart = i
            depth = depth + 1
        ElseIf ch = "}" Then
            depth = depth - 1
            If depth = 0 And entryStart > 0 Then
                result.Add Mid(json, entryStart, i - entryStart + 1)
                entryStart = 0
            End If
        ElseIf ch = "]" And depth = 0 Then
            Exit For
        End If
    Next i
    Set SplitMsgArray = result
End Function

' 從單一物件字串取巢狀 "trade":{"...","z":"<price>",...} 的 z 值。
' 頂層 "z" 刻意不用——兩次查詢之間常變回 "-"，trade.z 才穩定。
Private Function ExtractTradeZ(entryJson As String) As Double
    On Error GoTo Bail
    Dim tradeMarker As String: tradeMarker = """trade"":{"
    Dim tp As Long: tp = InStr(entryJson, tradeMarker)
    If tp = 0 Then GoTo Bail

    Dim tEnd As Long: tEnd = InStr(tp, entryJson, "}")
    If tEnd = 0 Then GoTo Bail
    Dim win As String: win = Mid(entryJson, tp, tEnd - tp + 1)

    Dim zMarker As String: zMarker = """z"":"""
    Dim zp As Long: zp = InStr(win, zMarker)
    If zp = 0 Then GoTo Bail

    Dim raw As String: raw = Mid(win, zp + Len(zMarker), 30)
    raw = Split(raw, """")(0)
    raw = Trim(Replace(raw, ",", ""))

    If IsNumeric(raw) Then
        Dim v As Double: v = CDbl(raw)
        If v > 0 Then ExtractTradeZ = v: Exit Function
    End If

Bail:
    ExtractTradeZ = 0
End Function

' 剝除 .TW / .TWO 後綴，取得純股票代號（輸入已 UCase）
Private Function StripSuffixMIS(ticker As String) As String
    Dim n As String: n = ticker
    If Len(n) >= 4 And Right(n, 4) = ".TWO" Then n = Left(n, Len(n) - 4)
    If Len(n) >= 3 And Right(n, 3) = ".TW" Then n = Left(n, Len(n) - 3)
    StripSuffixMIS = n
End Function

Private Function MisHttpGet(url As String) As String
    On Error GoTo Bail
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    On Error Resume Next
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.setRequestHeader "Accept", "application/json"
    http.send
    Dim ok As Boolean: ok = (Err.Number = 0)
    On Error GoTo Bail
    If ok And http.status = 200 Then MisHttpGet = http.responseText: Exit Function
Bail:
    MisHttpGet = ""
End Function
