Attribute VB_Name = "modLiveRefresh"
Option Explicit

' ================================================================
'  LIVE PRICE AUTO-REFRESH  (2026-09-23)
'
'  Public Sub StartDwellTimer()   - call from RR4's Worksheet_Activate
'  Public Sub StopAutoRefresh()   - call from RR4's Worksheet_Deactivate
'                                    and from Workbook_BeforeClose
'  Public Sub ArmAutoRefresh()    - Application.OnTime target, not for
'                                    direct use
'  Public Sub Tick()              - Application.OnTime target, not for
'                                    direct use
'
'  Design: stay on the RR4 ("P") sheet for DWELL_SEC continuously and a
'  TICK_SEC-interval refresh loop starts, updating just the LAST/price
'  cells of whichever market (TW/US) is open right now - it deliberately
'  does NOT run RebuildPortfolioDashboard (measured ~22s per run, and it
'  fully re-clears/redraws the sheet - wrong tool for a 15s tick). See
'  PortfolioDashboard_v3.RefreshLivePricesLite for the actual cell writes.
'  Leaving the P sheet (or closing the workbook) cancels whatever OnTime
'  call is pending; arriving back on P always restarts the 60s dwell from
'  scratch, it never resumes a partial wait.
' ================================================================

Private Const DWELL_SEC As Long = 60
Private Const TICK_SEC  As Long = 15
Private Const RR4_SHEET_NAME As String = "RR4"

Private g_scheduledAt   As Date     ' 0 = nothing pending
Private g_scheduledProc As String

Public Sub StartDwellTimer()
    Call CancelPending
    g_scheduledProc = "ArmAutoRefresh"
    g_scheduledAt = Now + TimeSerial(0, 0, DWELL_SEC)
    Application.OnTime g_scheduledAt, g_scheduledProc
End Sub

Public Sub StopAutoRefresh()
    Call CancelPending
End Sub

Private Sub CancelPending()
    If g_scheduledAt = 0 Then Exit Sub
    On Error Resume Next
    Application.OnTime EarliestTime:=g_scheduledAt, Procedure:=g_scheduledProc, Schedule:=False
    On Error GoTo 0
    g_scheduledAt = 0
    g_scheduledProc = ""
End Sub

' Fires once, DWELL_SEC after Worksheet_Activate armed it. If the user is
' still on P (Deactivate may have raced an OnTime call already in flight -
' this is the belt-and-suspenders check for that), kick off the tick loop.
Public Sub ArmAutoRefresh()
    g_scheduledAt = 0: g_scheduledProc = ""
    If Not RunningOnRR4() Then Exit Sub
    Call Tick
End Sub

Public Sub Tick()
    If Not RunningOnRR4() Then Exit Sub
    On Error Resume Next
    Call PortfolioDashboard_v3.RefreshLivePricesLite
    On Error GoTo 0

    g_scheduledProc = "Tick"
    g_scheduledAt = Now + TimeSerial(0, 0, TICK_SEC)
    Application.OnTime g_scheduledAt, g_scheduledProc
End Sub

Private Function RunningOnRR4() As Boolean
    On Error GoTo Bail
    RunningOnRR4 = (ActiveSheet.Name = RR4_SHEET_NAME)
    Exit Function
Bail:
    RunningOnRR4 = False
End Function
