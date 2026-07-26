Attribute VB_Name = "frmTransaction_Installer"
Option Explicit

' =============================================================================
' frmTransaction_Installer.bas
'
' Automatically builds the frmTransaction UserForm with all controls and code.
'
' HOW TO USE:
'   1. Alt+F11  -> File -> Import File -> select this .bas file -> OK
'   2. IMPORTANT: enable "Trust access to VBA project object model":
'      File -> Options -> Trust Center -> Trust Center Settings
'      -> Macro Settings -> check "Trust access to the VBA project object model"
'   3. Alt+F8 -> BuildFrmTransaction -> Run
'   4. The form is created. Test: Alt+F8 -> RR5_Core.ShowTransactionForm -> Run
'   5. Optional: delete this installer module (right-click -> Remove Module)
'
' Requirements: RR5_Core module must already be imported.
' frmTransaction_Code.txt must be in the same folder as the .xlsm file.
' =============================================================================

Public Sub BuildFrmTransaction()
    On Error GoTo ErrHandler

    Dim vbp As Object
    Set vbp = ThisWorkbook.VBProject

    ' Remove existing frmTransaction if present
    Dim comp As Object
    For Each comp In vbp.VBComponents
        If comp.Name = "frmTransaction" Then
            vbp.VBComponents.Remove comp
            Exit For
        End If
    Next comp

    ' Create new UserForm (vbext_ct_MSForm = 3)
    Dim uf As Object
    Set uf = vbp.VBComponents.Add(3)
    uf.Name = "frmTransaction"

    Dim d As Object
    Set d = uf.Designer

    ' =========================================================================
    ' Form properties
    ' =========================================================================
    d.Caption  = "Add Transaction - RR5"
    d.Width    = 420
    d.Height   = 580

    ' =========================================================================
    ' Controls  -- Left column (x~10, labels + inputs stacked)
    ' =========================================================================
    AddLbl d, "lblDate",       "Date:",              10,  10, 100, 14
    AddTxt d, "txtDate",                             10,  26, 130, 20
    AddLbl d, "lblTicker",     "Ticker:",            10,  52, 100, 14
    AddTxt d, "txtTicker",                           10,  68, 130, 20
    AddLbl d, "lblAction",     "Action:",            10,  94, 100, 14
    AddCmb d, "cmbAction",                           10, 110, 130, 20
    AddLbl d, "lblType",       "Type:",              10, 136, 100, 14
    AddCmb d, "cmbType",                             10, 152, 130, 20
    AddLbl d, "lblStrike",     "Strike:",            10, 178, 100, 14
    AddTxt d, "txtStrike",                           10, 194, 130, 20
    AddLbl d, "lblExpiry",     "Expiry:",            10, 220, 100, 14
    AddTxt d, "txtExpiry",                           10, 236, 130, 20

    ' =========================================================================
    ' Controls -- Right column (x~220)
    ' =========================================================================
    AddLbl d, "lblShares",     "Shares/Contracts:", 220,  10, 150, 14
    AddTxt d, "txtShares",                          220,  26, 150, 20
    AddLbl d, "lblMultiplier", "Multiplier:",       220,  52, 100, 14
    AddTxt d, "txtMultiplier",                      220,  68, 150, 20
    AddLbl d, "lblPrice",      "Price:",            220,  94, 100, 14
    AddTxt d, "txtPrice",                           220, 110, 150, 20
    AddLbl d, "lblRFR",        "Risk-Free Rate:",   220, 136, 130, 14
    AddTxt d, "txtRiskFreeRate",                    220, 152, 150, 20
    AddLbl d, "lblFee",        "Fee:",              220, 178, 100, 14
    AddTxt d, "txtFee",                             220, 194, 150, 20
    AddLbl d, "lblTax",        "Tax:",              220, 220, 100, 14
    AddTxt d, "txtTax",                             220, 236, 150, 20
    AddLbl d, "lblNetAmt",     "Net Amount (auto):",220, 262, 150, 14

    ' txtNetAmount -- read-only, dark background
    Dim c As Object
    Set c = d.Controls.Add("Forms.TextBox.1", "txtNetAmount", True)
    c.Left = 220: c.Top = 278: c.Width = 150: c.Height = 20
    c.Locked    = True
    c.BackColor = &H404040
    c.ForeColor = &HFFD700

    ' =========================================================================
    ' Controls -- Bottom section (y~270+)
    ' =========================================================================
    AddLbl d, "lblSector",    "Sector:",      10, 270,  80, 14
    AddTxt d, "txtSector",                   90, 268, 160, 20
    AddLbl d, "lblTarget",    "Target:",      10, 294,  80, 14
    AddTxt d, "txtTarget",                   90, 292, 160, 20
    AddLbl d, "lblStrategy",  "Strategy:",   10, 318,  80, 14
    AddCmb d, "cmbStrategy",                 90, 316, 160, 20
    AddLbl d, "lblIV",        "IV Entry:",   10, 342,  80, 14
    AddTxt d, "txtIV",                       90, 340, 100, 20
    AddLbl d, "lblDelta",     "Delta:",     200, 342,  50, 14
    AddTxt d, "txtDelta",                   250, 340, 120, 20
    AddLbl d, "lblBeta",      "Beta:",       10, 366,  80, 14
    AddTxt d, "txtBeta",                     90, 364, 100, 20
    AddLbl d, "lblBroker",    "Broker:",    10, 390,  80, 14
    AddCmb d, "cmbBroker",                  90, 388, 160, 20
    AddLbl d, "lblUnderlying","Underlying:", 10, 414,  80, 14
    AddTxt d, "txtUnderlying",              90, 412, 160, 20
    AddLbl d, "lblUndlSpot",  "Spot Px:",  260, 414,  50, 14
    AddTxt d, "txtUndlSpot",                310, 412,  60, 20

    ' Conv Ratio -- required for warrants; label starts grey, turns orange via cmbType_Change
    Set c = d.Controls.Add("Forms.Label.1", "lblConvRatio", True)
    c.Caption = "Conv Ratio:": c.Left = 10: c.Top = 438: c.Width = 80: c.Height = 14
    c.ForeColor = &H606060   ' grey by default (not a warrant yet)
    c.BackStyle = 0
    AddTxt d, "txtConvRatio", 90, 436, 100, 20

    Set c = d.Controls.Add("Forms.Label.1", "lblConvRatioHint", True)
    c.Caption = "(e.g. 0.1 = 10 warrants per share)"
    c.Left = 196: c.Top = 440: c.Width = 200: c.Height = 12
    c.ForeColor = &H808080: c.BackStyle = 0
    c.Font.Size = 7

    ' =========================================================================
    ' Buttons (bottom strip)
    ' =========================================================================
    Set c = d.Controls.Add("Forms.CommandButton.1", "btnSave", True)
    c.Caption = "Save":   c.Left = 20:  c.Top = 468: c.Width = 110: c.Height = 26
    c.Default = True
    c.BackColor = &H6400    ' dark green
    c.ForeColor = &HFFFFFF

    Set c = d.Controls.Add("Forms.CommandButton.1", "btnClear", True)
    c.Caption = "Clear":  c.Left = 150: c.Top = 468: c.Width = 100: c.Height = 26

    Set c = d.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    c.Caption = "Cancel": c.Left = 270: c.Top = 468: c.Width = 110: c.Height = 26
    c.Cancel = True

    ' =========================================================================
    ' Insert form VBA code
    ' =========================================================================
    InsertFormCode uf

    MsgBox "frmTransaction built successfully!" & vbCrLf & vbCrLf & _
           "Test: Alt+F8 -> RR5_Core.ShowTransactionForm -> Run" & vbCrLf & vbCrLf & _
           "You may now delete this installer module.", _
           vbInformation, "RR5 Installer"
    Exit Sub

ErrHandler:
    If Err.Number = 50289 Then
        MsgBox "ERROR: VBA project access not trusted." & vbCrLf & vbCrLf & _
               "Enable it first:" & vbCrLf & _
               "File -> Options -> Trust Center -> Trust Center Settings" & vbCrLf & _
               "-> Macro Settings -> check" & vbCrLf & _
               """Trust access to the VBA project object model""" & vbCrLf & vbCrLf & _
               "Then re-run BuildFrmTransaction.", vbCritical, "RR5 Installer"
    Else
        MsgBox "Build failed: " & Err.Description & " (Err " & Err.Number & ")", _
               vbCritical, "RR5 Installer"
    End If
End Sub

' =============================================================================
' Private helpers
' =============================================================================

Private Sub AddLbl(d As Object, nm As String, cap As String, _
                   x As Long, y As Long, w As Long, h As Long)
    Dim c As Object
    Set c = d.Controls.Add("Forms.Label.1", nm, True)
    c.Caption   = cap
    c.Left = x: c.Top = y: c.Width = w: c.Height = h
    c.ForeColor = &H00B050   ' green
    c.BackStyle = 0          ' transparent
End Sub

Private Sub AddTxt(d As Object, nm As String, _
                   x As Long, y As Long, w As Long, h As Long)
    Dim c As Object
    Set c = d.Controls.Add("Forms.TextBox.1", nm, True)
    c.Left = x: c.Top = y: c.Width = w: c.Height = h
    c.BackColor = &H1A1A1A
    c.ForeColor = &H87CEEB   ' light blue
End Sub

Private Sub AddCmb(d As Object, nm As String, _
                   x As Long, y As Long, w As Long, h As Long)
    Dim c As Object
    Set c = d.Controls.Add("Forms.ComboBox.1", nm, True)
    c.Left = x: c.Top = y: c.Width = w: c.Height = h
    c.BackColor = &H1A1A1A
    c.ForeColor = &H87CEEB
    c.Style = 0              ' fmStyleDropDownCombo (editable + dropdown arrow)
End Sub

' =============================================================================
' InsertFormCode: reads frmTransaction_Code.txt from ThisWorkbook.Path
'                and inserts it into the UserForm code module.
' =============================================================================
Private Sub InsertFormCode(uf As Object)
    Dim fPath As String
    fPath = ThisWorkbook.Path & "\frmTransaction_Code.txt"

    If Dir(fPath) = "" Then
        MsgBox "frmTransaction_Code.txt not found at:" & vbCrLf & fPath & vbCrLf & vbCrLf & _
               "Form controls created, but code not inserted." & vbCrLf & _
               "Paste frmTransaction_Code.txt manually into the form's code window.", _
               vbExclamation, "RR5 Installer"
        Exit Sub
    End If

    ' Read the file
    Dim fNo As Integer: fNo = FreeFile
    Dim allCode As String, oneLine As String
    Open fPath For Input As #fNo
    Do While Not EOF(fNo)
        Line Input #fNo, oneLine
        allCode = allCode & oneLine & vbCrLf
    Loop
    Close #fNo

    ' Insert into form code module
    Dim cm As Object
    Set cm = uf.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.InsertLines 1, allCode
End Sub
