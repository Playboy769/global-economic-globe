Attribute VB_Name = "TaiwanPriceFetcher"
Option Explicit

' ================================================================
'  TAIWAN PRICE FETCHER  v1.1
'
'  Public Function GetTWStockPrice(ticker As String) As Double
'  Public Sub     ClearPriceCache()
'
'  專門處理台股（.TW / .TWO），不碰美股。
'  美股仍走 Attach 模組的 GetStockPrice（Yahoo Finance）。
'  兩模組 public 名稱不重疊，無衝突。
'
'  TickerInsight.bas 的 GetStockPriceSafe 負責路由：
'    .TW / .TWO suffix -> GetTWStockPrice (此模組)
'    其餘              -> GetStockPrice   (Attach 模組)
'
'  Session 快取:
'    STOCK_DAY_ALL 和 tpex_mainboard_daily_close_quotes 每次工作簿
'    開啟後只抓一次，同一個 session 內所有台股 ticker 共用快取。
'    Call ClearPriceCache() 可強制重新抓取（例如跨日後第一次查詢）。
' ================================================================

' ----- API 端點 -----
Private Const TWSE_URL As String = _
    "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"

Private Const TPEX_URL As String = _
    "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes"

' JSON 欄位名稱
' TWSE STOCK_DAY_ALL          -> Code / ClosingPrice
' TPEx mainboard_close_quotes -> SecuritiesCode / ClosingPrice
Private Const TWSE_CODE_KEY As String = "Code"
Private Const TPEX_CODE_KEY As String = "SecuritiesCode"
Private Const PRICE_KEY     As String = "ClosingPrice"

Private Const HTTP_RETRY   As Integer = 3
Private Const HTTP_WAIT_MS As Long = 800

' ----- Session 快取 (模組層級，工作簿關閉前持續有效) -----
Private g_twseJson  As String
Private g_tpexJson  As String
Private g_twseReady As Boolean
Private g_tpexReady As Boolean

' ================================================================
'  PUBLIC ENTRY POINT — 台股專用，非台股回傳 0
' ================================================================
Public Function GetTWStockPrice(ticker As String) As Double
    Dim t As String: t = UCase(Trim(CStr(ticker)))
    If Len(t) = 0 Then GetTWStockPrice = 0: Exit Function

    Dim code As String: code = StripSuffix(t)

    Select Case True
        Case InStr(t, ".TWO") > 0:  GetTWStockPrice = TPExPrice(code)
        Case InStr(t, ".TW") > 0:   GetTWStockPrice = TWSEPrice(code)
        Case Else:                   GetTWStockPrice = 0  ' 非台股請走 Attach 的 GetStockPrice
    End Select
End Function

' 強制清除快取，下次查詢時重新抓取
Public Sub ClearPriceCache()
    g_twseJson = "":  g_twseReady = False
    g_tpexJson = "":  g_tpexReady = False
End Sub

' ================================================================
'  TWSE (上市) — STOCK_DAY_ALL
'  JSON 格式: [{"Code":"2330","ClosingPrice":"900.00",...}, ...]
' ================================================================
Private Function TWSEPrice(code As String) As Double
    On Error GoTo Bail
    If Not g_twseReady Then
        Application.StatusBar = "TWSE: 抓取上市收盤行情..."
        g_twseJson = HttpGet(TWSE_URL)
        g_twseReady = (LenB(g_twseJson) > 0)   ' 只有成功取到資料才標 ready，失敗時允許下次重試
        Application.StatusBar = False
    End If
    If LenB(g_twseJson) > 0 Then
        TWSEPrice = ParsePrice(g_twseJson, TWSE_CODE_KEY, code, PRICE_KEY)
    End If
    Exit Function
Bail:
    TWSEPrice = 0
End Function

' ================================================================
'  TPEx (上櫃) — tpex_mainboard_daily_close_quotes
'  JSON 格式: [{"SecuritiesCode":"6547","ClosingPrice":"xx.xx",...}, ...]
' ================================================================
Private Function TPExPrice(code As String) As Double
    On Error GoTo Bail
    If Not g_tpexReady Then
        Application.StatusBar = "TPEx: 抓取上櫃收盤行情..."
        g_tpexJson = HttpGet(TPEX_URL)
        g_tpexReady = (LenB(g_tpexJson) > 0)   ' 只有成功取到資料才標 ready，失敗時允許下次重試
        Application.StatusBar = False
    End If
    If LenB(g_tpexJson) > 0 Then
        TPExPrice = ParsePrice(g_tpexJson, TPEX_CODE_KEY, code, PRICE_KEY)
    End If
    Exit Function
Bail:
    TPExPrice = 0
End Function

' ================================================================
'  JSON 解析器（適用於扁平陣列物件）
'
'  邏輯:
'  1. 在 JSON 中找到 "keyField":"code"
'  2. 從命中位置向前搜尋到當筆記錄結尾 "}"（避免抓到上一筆的欄位）
'  3. 在視窗內找到 "valField" 並解析數值（支援字串或數字欄位）
' ================================================================
Private Function ParsePrice(json As String, _
                             keyField As String, code As String, _
                             valField As String) As Double
    On Error GoTo Bail

    Dim needle As String: needle = """" & keyField & """:""" & code & """"
    Dim keyPos As Long:   keyPos = InStr(json, needle)
    If keyPos = 0 Then GoTo Bail

    ' 從 key 命中位置向後取到本筆記錄結尾 }（TWSE/TPEx 為簡單扁平物件，無巢狀）
    Dim afterKey As Long: afterKey = keyPos + Len(needle)
    Dim recEnd As Long:   recEnd = InStr(afterKey, json, "}")
    If recEnd = 0 Or recEnd > afterKey + 500 Then recEnd = afterKey + 500
    Dim win As String:    win = Mid(json, afterKey, recEnd - afterKey + 1)

    ' 嘗試字串值欄位: "valField":"<number>"
    Dim vfStr As String: vfStr = """" & valField & """:"""
    Dim vp As Long:      vp = InStr(win, vfStr)
    Dim raw As String

    If vp > 0 Then
        raw = Mid(win, vp + Len(vfStr), 30)
        raw = Split(raw, """")(0)                    ' 取到結尾 "
    Else
        ' 嘗試數字值欄位: "valField":<number>
        vfStr = """" & valField & """:"
        vp = InStr(win, vfStr)
        If vp = 0 Then GoTo Bail
        raw = Mid(win, vp + Len(vfStr), 30)
        raw = Split(Split(raw, ",")(0), "}")(0)      ' 取到 , 或 }
    End If

    raw = Trim(Replace(raw, ",", ""))                ' 移除千分位符號
    If IsNumeric(raw) And CDbl(raw) > 0 Then
        ParsePrice = CDbl(raw)
        Exit Function                                ' ← 必須在 Bail 前離開
    End If

Bail:
    ParsePrice = 0
End Function

' ================================================================
'  工具函式
' ================================================================

' 剝除 .TW / .TWO 後綴，取得純股票代號（輸入已 UCase）
Private Function StripSuffix(ticker As String) As String
    Dim n As String: n = ticker
    If Len(n) >= 4 And Right(n, 4) = ".TWO" Then n = Left(n, Len(n) - 4)
    If Len(n) >= 3 And Right(n, 3) = ".TW"  Then n = Left(n, Len(n) - 3)
    StripSuffix = n
End Function

Private Function HttpGet(url As String) As String
    On Error GoTo Bail
    Dim attempt As Integer
    For attempt = 1 To HTTP_RETRY
        Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
        On Error Resume Next
        http.Open "GET", url, False
        http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0)"
        http.setRequestHeader "Accept",     "application/json"
        http.send
        Dim ok As Boolean: ok = (Err.Number = 0)
        On Error GoTo Bail
        If ok And http.status = 200 Then HttpGet = http.responseText: Exit Function
        Set http = Nothing
        On Error GoTo Bail
        If attempt < HTTP_RETRY Then PauseMS HTTP_WAIT_MS * attempt
    Next attempt
Bail:
    HttpGet = ""
End Function

Private Sub PauseMS(ms As Long)
    Dim t As Single: t = Timer + ms / 1000
    Do While Timer < t: DoEvents: Loop
End Sub
