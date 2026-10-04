Attribute VB_Name = "modHttp"
Option Explicit

' SEC requires every automated request to carry an identifying User-Agent
' (name + contact email) -- https://www.sec.gov/os/webmaster-faq#developers
Public Const SEC_USER_AGENT As String = "Ryan Personal Research Tool ryan929929@gmail.com"

Private Const HTTP_MAX_ATTEMPTS As Long = 3

' WinHttpRequest timeouts, in milliseconds: (resolve, connect, send, receive).
' Without these, a single stalled request (dead connection, SEC hanging under
' load) blocks Excel indefinitely with no way to recover short of killing the
' process -- especially risky now that HttpGet retries up to 3x. Receive is
' the generous one since a large company's companyfacts JSON can be several
' MB on a slow connection; resolve/connect/send failing that slowly would mean
' something is genuinely wrong, not just "slow server".
Private Const HTTP_TIMEOUT_RESOLVE_MS As Long = 10000
Private Const HTTP_TIMEOUT_CONNECT_MS As Long = 10000
Private Const HTTP_TIMEOUT_SEND_MS As Long = 15000
Private Const HTTP_TIMEOUT_RECEIVE_MS As Long = 60000

' Performs a GET request and returns the response body as text.
' Retries on rate-limiting (429), server errors (5xx), and network-level
' failures (timeout/DNS/connection reset), with a short linear backoff (1s,
' then 2s) between attempts. A 404 or other 4xx fails immediately without
' retrying -- those mean the resource genuinely doesn't exist (e.g. a bad
' CIK), so waiting and retrying would just waste time.
' Raises an error (via Err.Raise) if the final attempt still isn't a 200.
Public Function HttpGet(ByVal url As String) As String
    Dim attempt As Long
    For attempt = 1 To HTTP_MAX_ATTEMPTS
        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.SetTimeouts HTTP_TIMEOUT_RESOLVE_MS, HTTP_TIMEOUT_CONNECT_MS, HTTP_TIMEOUT_SEND_MS, HTTP_TIMEOUT_RECEIVE_MS

        Dim comErrNum As Long, comErrDesc As String
        On Error Resume Next
        http.Open "GET", url, False
        http.SetRequestHeader "User-Agent", SEC_USER_AGENT
        http.SetRequestHeader "Accept", "application/json, text/html, */*"
        http.Send
        comErrNum = Err.Number
        comErrDesc = Err.Description
        On Error GoTo 0

        If comErrNum <> 0 Then
            If attempt = HTTP_MAX_ATTEMPTS Then
                Err.Raise vbObjectError + 5002, "HttpGet", "網路錯誤，重試 " & HTTP_MAX_ATTEMPTS & " 次後仍失敗：" & comErrDesc & " for " & url
            End If
        ElseIf http.Status = 200 Then
            HttpGet = Utf8Body(http)
            Exit Function
        ElseIf http.Status = 429 Or http.Status >= 500 Then
            If attempt = HTTP_MAX_ATTEMPTS Then
                Err.Raise vbObjectError + 5001, "HttpGet", "HTTP " & http.Status & " " & http.StatusText & " 重試 " & HTTP_MAX_ATTEMPTS & " 次後仍失敗 for " & url
            End If
        Else
            Err.Raise vbObjectError + 5001, "HttpGet", "HTTP " & http.Status & " " & http.StatusText & " for " & url
        End If

        Application.Wait Now + TimeSerial(0, 0, attempt)
    Next attempt
End Function

' The response body decoded as UTF-8. SEC (and Yahoo) serve UTF-8 JSON as
' "application/json" with no charset, and WinHttp's ResponseText then decodes
' the bytes with the wrong code page: SEC's entityName
' "KRATOS DEFENSE<U+00A0>& SECURITY SOLUTIONS,<U+00A0>INC." came out with an
' A-circumflex before every such space. MOPS pages are Big5 and go through
' modMOPSData.HttpGetBig5, not HttpGet. Falls back to ResponseText if the
' stream decode fails (e.g. an empty body).
Private Function Utf8Body(ByVal http As Object) As String
    On Error GoTo Fallback
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                      ' adTypeBinary
    stm.Open
    stm.Write http.ResponseBody
    stm.Position = 0
    stm.Type = 2                      ' adTypeText
    stm.Charset = "utf-8"
    Utf8Body = stm.ReadText
    stm.Close
    Exit Function
Fallback:
    Utf8Body = http.ResponseText
End Function
