Attribute VB_Name = "modLivePrice"
Option Explicit

' ================================================================
'  LIVE PRICE FETCHER — used ONLY by the 15s auto-refresh tick
'  (PortfolioDashboard_v3.RefreshLivePricesLite)  (2026-09-23)
'
'  Public Function GetLivePrice(ticker As String) As Double
'
'  Deliberately independent of Attach.GetStockPrice: that path carries
'  its own 5-minute g_PriceCache, and even after adding a ForceFresh
'  bypass flag to it, the auto-refresh tick still kept showing frozen
'  prices in practice. Rather than keep chasing where staleness was
'  creeping back in through a function shared by half a dozen other
'  callers, this module has NO caching of any kind anywhere in it -
'  every single call is a genuinely uncached read, guaranteed by
'  construction, not by a flag someone could forget to pass:
'    .TW / .TWO  -> modMISPrice.GetMISPrice (the per-tick batch
'                   prefetch RefreshLivePricesLite already runs before
'                   this is called - that part was screenshot-verified
'                   fresh earlier and was never the suspect)
'    everything else -> direct Yahoo chart API hit, every call, no
'                   memoization, same parsing logic as Attach.bas's
'                   ParseClosePrice (duplicated rather than shared, so
'                   this module truly touches nothing in Attach.bas)
' ================================================================

Public Function GetLivePrice(ByVal ticker As String) As Double
    GetLivePrice = 0
    Dim t As String: t = UCase(Trim(ticker))
    If Len(t) = 0 Then Exit Function

    If InStr(t, ".TW") > 0 Then
        GetLivePrice = modMISPrice.GetMISPrice(t)
        Exit Function
    End If

    On Error GoTo Bail
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim safeTicker As String: safeTicker = Replace(t, "^", "%5E")
    Dim url As String
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & safeTicker & "?interval=1d&range=5d"

    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.setRequestHeader "Accept", "application/json"
    http.send
    GetLivePrice = ParseLiveClose(http.responseText)
    Exit Function

Bail:
    GetLivePrice = 0
End Function

' Same field priority as Attach.ParseClosePrice: regularMarketPrice
' (near-real-time quote) first, the daily close series as fallback.
Private Function ParseLiveClose(ByVal resp As String) As Double
    ParseLiveClose = 0

    Dim metaPos As Long: metaPos = InStr(resp, """regularMarketPrice"":")
    If metaPos > 0 Then
        Dim vals As Long: vals = metaPos + 21
        Dim valE As Long: valE = InStr(vals, resp, ",")
        If valE = 0 Then valE = InStr(vals, resp, "}")
        If valE > vals Then
            Dim metaVal As String: metaVal = Trim(Mid(resp, vals, valE - vals))
            If IsNumeric(metaVal) And Val(metaVal) > 0 Then
                ParseLiveClose = Val(metaVal)
                Exit Function
            End If
        End If
    End If

    Dim csPos As Long: csPos = InStr(resp, """close"":[")
    If csPos = 0 Then Exit Function
    Dim csS As Long: csS = csPos + 9
    Dim csE As Long: csE = InStr(csS, resp, "]")
    If csE <= csS Then Exit Function

    Dim parts() As String: parts = Split(Mid(resp, csS, csE - csS), ",")
    Dim i As Long
    For i = UBound(parts) To 0 Step -1
        Dim pt As String: pt = Trim(parts(i))
        If IsNumeric(pt) Then
            If CDbl(pt) > 0 Then
                ParseLiveClose = CDbl(pt)
                Exit Function
            End If
        End If
    Next i
End Function
