Attribute VB_Name = "CompanyResearchSEC"
Option Explicit

' ============================================================================
'  CompanyResearchSEC  --  RR4 "Company research" sheet, ticker hand-off
' ----------------------------------------------------------------------------
'  2026-09-13: the single-ticker financial deep-dive that used to render in
'  this sheet's lower half (rows 40+) moved to its own page, "Earnings"
'  (nav code E, RR4/modEarnings.bas).  This module now only forwards:
'
'    - typing a ticker into D2 (or the F2 market override) on Company research
'      -> RunDeepDive -> modEarnings.ShowEarnings(ticker, market) + jump to E
'    - double-clicking a ticker in the scan table -> RunDeepDiveFromScan
'
'  RunDeepDive also blanks rows 40+ once, so a sheet that still carries an old
'  deep-dive loses it on the first hand-off.  The two public names and the two
'  input-cell constants are kept because RR4/SheetCompanyResearch_Code.txt and
'  Sanner.bas refer to them.
'  Pure ASCII (RR4 .bas convention).
' ============================================================================

Public Const CR_SHEET       As String = "Company research"
Public Const CR_INPUT_CELL  As String = "D2"    ' CR page v4: the strip is horizontal on row 2
Public Const CR_MARKET_CELL As String = "F2"

Private Const OLD_BAND_ROW   As Long = 40       ' first row of the former deep-dive band
Private Const CLEAR_LAST_ROW As Long = 400
Private Const CLR_TEXT       As Long = 15132390 ' RGB(230,230,230)
Private Const FONT_FACE      As String = "Consolas"

' ---------------------------------------------------------------------------
'  Entry point (Worksheet_Change on D2 / F2)
' ---------------------------------------------------------------------------
Public Sub RunDeepDive(ByVal rawTicker As String, Optional ByVal marketOverride As String = "")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CR_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then Call ClearOldBand(ws)

    Dim tk As String
    tk = Trim$(rawTicker)
    If tk = "" Then Exit Sub
    Call modEarnings.ShowEarnings(tk, marketOverride, True)
End Sub

' Rows 40+ back to plain black - the deep-dive no longer lives here.
Private Sub ClearOldBand(ByVal ws As Worksheet)
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(OLD_BAND_ROW, 1), ws.Cells(CLEAR_LAST_ROW, 60))) > 0 Then
        With ws.Range(ws.Cells(OLD_BAND_ROW, 1), ws.Cells(CLEAR_LAST_ROW, 60))
            .Clear
            .Interior.Color = 0
            .Font.Color = CLR_TEXT
            .Font.Name = FONT_FACE
        End With
    End If
    Application.EnableEvents = prevEv
End Sub

' Double-click on a ticker in the scan table -> Earnings page for it.
' Called from the sheet's Worksheet_BeforeDoubleClick; True = handled.
Public Function RunDeepDiveFromScan(ByVal Target As Range) As Boolean
    If Target.Column <> Sanner.SC_TICKER Then Exit Function
    If Target.row < Sanner.SCAN_FIRST_ROW Or Target.row > Sanner.SCAN_LAST_ROW Then Exit Function
    Dim tk As String: tk = Trim$(CStr(Target.Value))
    If Len(tk) = 0 Or UCase$(tk) = "SUMMARY" Then Exit Function
    RunDeepDiveFromScan = True
    ' writing the input cell fires Worksheet_Change -> RunDeepDive
    Target.Worksheet.Range(CR_INPUT_CELL).Value = tk
End Function
