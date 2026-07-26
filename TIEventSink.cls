VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TIEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' TIEventSink - Application-level event router for Ticker Insight
'
' Why this class exists:
'   Workbook_SheetChange (in ThisWorkbook) is the "official" event hook
'   but requires the user to manually paste code into ThisWorkbook's
'   code module. Many users can't / don't, so B1 typing never triggers.
'
'   This class uses WithEvents on Application to hook the same event
'   purely from a regular class module. No ThisWorkbook edits needed.
'
' How it gets activated:
'   TickerInsight.bas declares  Public g_TISink As TIEventSink
'   OpenTickerInsight() calls   Set g_TISink = New TIEventSink
'   From that point on, every workbook change in this Excel session
'   is routed through App_SheetChange below.
'
'   The sink persists for the workbook session. Closing & reopening
'   Excel requires running OpenTickerInsight (or InstallTickerInsightEvents)
'   once again. Add InstallTickerInsightEvents to ThisWorkbook.Workbook_Open
'   if you want it persistent across opens.
' ================================================================

Private WithEvents App As Application

Private Sub Class_Initialize()
    Set App = Application
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub

' Fires on EVERY sheet change in EVERY workbook for this Excel instance.
' We filter to Ticker Insight only, and only inside ThisWorkbook (so a
' second workbook with a same-named sheet doesn't get hijacked).
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next

    ' Scope guard: only react to our own workbook's sheet
    If Sh.Parent Is Nothing Then Exit Sub
    If Not Sh.Parent Is ThisWorkbook Then Exit Sub

    ' Name guard
    If Sh.Name <> "Ticker Insight" Then Exit Sub

    Call TickerInsight_OnChange(Target)
End Sub
