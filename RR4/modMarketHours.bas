Attribute VB_Name = "modMarketHours"
Option Explicit

' ================================================================
'  MARKET HOURS  (2026-09-23)
'
'  Public Function IsTWMarketOpen() As Boolean
'  Public Function IsUSMarketOpen() As Boolean
'
'  Pure time-window check, no holiday calendar - a national holiday just
'  means a few pointless queries that return the same closing price, not
'  a functional bug (this is the same tradeoff sector-rotation-system's
'  own scheduler makes). US window uses a fixed ET offset with no DST
'  adjustment, per user: Taipei = ET + 12h year-round.
' ================================================================

Public Function IsTWMarketOpen() As Boolean
    Dim n As Date: n = Now
    If Weekday(n, vbMonday) > 5 Then IsTWMarketOpen = False: Exit Function
    Dim t As Date: t = TimeValue(n)
    IsTWMarketOpen = (t >= TimeSerial(9, 0, 0) And t <= TimeSerial(13, 30, 0))
End Function

Public Function IsUSMarketOpen() As Boolean
    Dim etNow As Date: etNow = DateAdd("h", -12, Now)   ' Taipei = ET + 12h, no DST
    If Weekday(etNow, vbMonday) > 5 Then IsUSMarketOpen = False: Exit Function
    Dim t As Date: t = TimeValue(etNow)
    IsUSMarketOpen = (t >= TimeSerial(9, 30, 0) And t <= TimeSerial(16, 0, 0))
End Function
