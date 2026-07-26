Attribute VB_Name = "modHttp"
Option Explicit

' SEC requires every automated request to carry an identifying User-Agent
' (name + contact email) -- https://www.sec.gov/os/webmaster-faq#developers
Public Const SEC_USER_AGENT As String = "Ryan Personal Research Tool ryan929929@gmail.com"

Private Const HTTP_MAX_ATTEMPTS As Long = 3

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
            HttpGet = http.ResponseText
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
