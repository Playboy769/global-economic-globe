VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTransaction 
   Caption         =   "UserForm1"
   ClientHeight    =   11077
   ClientLeft      =   99
   ClientTop       =   429
   ClientWidth     =   8206.001
   OleObjectBlob   =   "frmTransaction.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================
' frmTransaction - UserForm code module
'
' Usage: After creating a new UserForm in VBA Editor and renaming it
'        to "frmTransaction", paste this ENTIRE block into the form's
'        code window (double-click the form's empty area to open it).
'
' Requirements:
'   - RR5_Core module imported (provides GenerateTxID + AppendTransaction)
'   - Transactions sheet exists with 24-column header layout
'   - Controls: txtDate txtTicker txtStrike txtExpiry txtShares txtMultiplier
'               txtPrice txtRiskFreeRate txtFee txtTax txtNetAmount
'               txtSector txtTarget txtIV txtDelta txtBeta
'               txtUnderlying txtUndlSpot txtGamma txtTheta
'               cmbAction cmbType cmbStrategy cmbBroker
'               btnSave btnClear btnCancel
' =============================================================

Option Explicit

' --- UserForm_Initialize: populate dropdowns and set defaults on open ---
Private Sub UserForm_Initialize()
    cmbAction.List = Split("Buy,Sell,AdjustCost,Expire,Exercise", ",")
    ' 2026-10-18: Warrant family (Call Warrant/Put Warrant/Bull Cert/Bear Cert)
    ' removed - see CLAUDE.md "RR5 Warrant 移除". Remaining order is already a
    ' grouping: options (Call/Put), then Future, then cash instruments.
    cmbType.List = Split("Call,Put,Future,Equity,ETF", ",")
    cmbStrategy.List = Split("Bull Call,Bear Put,Straddle,Strangle,Iron Condor,Protective,Covered,Calendar,Ratio,Single", ",")
    cmbBroker.List = Split("Fubon,SinoPac,Cathay,KGI,Yuanta,Capital", ",")

    txtDate.Value = Format(Date, "yyyy/mm/dd")
    txtFee.Value = "0"
    txtTax.Value = "0"
    txtMultiplier.Value = "1"

    ' Load current risk-free rate from _Config (drives Black-Scholes Greeks).
    ' Stored as decimal fraction in _Config; e.g. 0.02 = 2 %.
    Dim curR As Variant
    curR = RR5_Core.GetConfigValue("BS_RISK_FREE_RATE")
    If IsNumeric(curR) And CDbl(curR) > 0 Then
        txtRiskFreeRate.Value = Format(CDbl(curR), "0.0000")
    Else
        txtRiskFreeRate.Value = "0.0200"   ' fallback default 2 %
    End If

    Me.Caption = "Add Transaction - RR5"
End Sub

' --- Auto-recalc Net Amount when any pricing field changes ---
Private Sub txtShares_Change():     RecalcNet: End Sub
Private Sub txtPrice_Change():      RecalcNet: End Sub
Private Sub txtMultiplier_Change(): RecalcNet: End Sub
Private Sub txtFee_Change():        RecalcNet: End Sub
Private Sub txtTax_Change():        RecalcNet: End Sub

Private Sub RecalcNet()
    Dim s As Double, p As Double, m As Double, f As Double, tx As Double
    s = val(txtShares.Value)
    p = val(txtPrice.Value)
    m = val(txtMultiplier.Value): If m = 0 Then m = 1
    f = val(txtFee.Value)
    tx = val(txtTax.Value)
    txtNetAmount.Value = Format(s * p * m + f + tx, "#,##0.00")
End Sub

' --- AutoFillEntryGreeks: silently compute Delta/Gamma/Theta at entry ---
' Only fires if (1) all 5 BS inputs are present, (2) the target field is blank
' (so user's manual entry always wins). Designed to be called on Save.
'
' Required inputs (all on form):
'   txtUndlSpot  -- underlying spot price at entry (must add as TextBox)
'   txtStrike    -- option strike
'   txtIV        -- implied volatility at entry (decimal, e.g. 0.18)
'   txtRiskFreeRate -- risk-free rate (decimal)
'   txtDate / txtExpiry -- for time to expiry
'   cmbType      -- "Call" or "Put"
Private Sub AutoFillEntryGreeks()
    On Error GoTo SafeExit   ' if txtUndlSpot doesn't exist on form yet, abort silently

    ' Only meaningful for options (Call / Put)
    Dim t As String: t = cmbType.Value
    If t <> "Call" And t <> "Put" Then Exit Sub

    ' All numeric inputs must be present and positive
    If Not IsNumeric(txtUndlSpot.Value) Then Exit Sub
    If Not IsNumeric(txtStrike.Value) Then Exit Sub
    If Not IsNumeric(txtIV.Value) Then Exit Sub
    If Not IsNumeric(txtRiskFreeRate.Value) Then Exit Sub

    Dim s As Double, k As Double, v As Double, r As Double
    s = CDbl(txtUndlSpot.Value)
    k = CDbl(txtStrike.Value)
    v = CDbl(txtIV.Value)
    r = CDbl(txtRiskFreeRate.Value)
    If s <= 0 Or k <= 0 Or v <= 0 Or r <= 0 Then Exit Sub

    ' Time to expiry in years
    If Not IsDate(txtDate.Value) Or Not IsDate(txtExpiry.Value) Then Exit Sub
    Dim T_years As Double
    T_years = (CDate(txtExpiry.Value) - CDate(txtDate.Value)) / 365
    If T_years <= 0 Then Exit Sub

    Dim bsType As String
    Select Case t
        Case "Call": bsType = "Call"
        Case Else:   bsType = "Put"
    End Select

    ' Only overwrite empty fields so user's manual entries win
    If Trim(CStr(txtDelta.Value)) = "" Then
        txtDelta.Value = Format(RR5_Core.BS_Delta(s, k, T_years, r, v, bsType), "0.0000")
    End If
    If Trim(CStr(txtGamma.Value)) = "" Then
        txtGamma.Value = Format(RR5_Core.BS_Gamma(s, k, T_years, r, v), "0.000000")
    End If
    If Trim(CStr(txtTheta.Value)) = "" Then
        txtTheta.Value = Format(RR5_Core.BS_Theta(s, k, T_years, r, v, bsType), "0.0000")
    End If
SafeExit:
End Sub

' --- Save button: validate, build row array, append to Transactions sheet ---
Private Sub btnSave_Click()
    ' Required field validation
    If Trim(txtTicker.Value) = "" Then
        MsgBox "Ticker required.", vbExclamation: txtTicker.SetFocus: Exit Sub
    End If
    If cmbAction.Value = "" Then
        MsgBox "Action required.", vbExclamation: cmbAction.SetFocus: Exit Sub
    End If
    If val(txtShares.Value) = 0 Then
        MsgBox "Shares/Contracts must be > 0.", vbExclamation: txtShares.SetFocus: Exit Sub
    End If
    If val(txtPrice.Value) = 0 Then
        MsgBox "Price must be > 0.", vbExclamation: txtPrice.SetFocus: Exit Sub
    End If

    ' Risk-Free Rate: validate (decimal, 0 < r < 1) and persist to _Config
    Dim rfr As Double
    rfr = val(txtRiskFreeRate.Value)
    If rfr <= 0 Or rfr >= 1 Then
        MsgBox "Risk-Free Rate must be a decimal between 0 and 1." & vbCrLf & _
               "Example: 0.02 = 2 %, 0.045 = 4.5 %", _
               vbExclamation: txtRiskFreeRate.SetFocus: Exit Sub
    End If
    Call RR5_Core.SetConfigValue("BS_RISK_FREE_RATE", rfr)

    ' Auto-fill entry-time Delta/Gamma/Theta if all BS inputs are present
    ' (silently skipped for Equity/ETF or if any input missing).
    Call AutoFillEntryGreeks

    ' Build 24-element row array matching Transactions sheet column order
    ' Col:  1=ID  2=Date  3=Ticker  4=Action  5=Type  6=Strike  7=Expiry
    '       8=Qty 9=Mult  10=Price  11=Fee    12=Tax  13=NetAmt
    '      14=Sector 15=Target 16=Strategy 17=IV 18=Delta 19=Beta 20=Broker
    '      21=Underlying 22=Gamma 23=Theta 24=Conv_Ratio
    Dim rowData(1 To 24) As Variant
    rowData(1) = RR5_Core.GenerateTxID()
    rowData(2) = SafeDate(txtDate.Value)
    rowData(3) = Trim(txtTicker.Value)
    rowData(4) = cmbAction.Value
    rowData(5) = cmbType.Value
    rowData(6) = SafeNum(txtStrike.Value)
    rowData(7) = SafeDate(txtExpiry.Value)
    rowData(8) = val(txtShares.Value)
    rowData(9) = IIf(val(txtMultiplier.Value) = 0, 1, val(txtMultiplier.Value))
    rowData(10) = val(txtPrice.Value)
    rowData(11) = val(txtFee.Value)
    rowData(12) = val(txtTax.Value)
    rowData(13) = val(Replace(txtNetAmount.Value, ",", ""))
    rowData(14) = txtSector.Value
    rowData(15) = SafeNum(txtTarget.Value)
    rowData(16) = cmbStrategy.Value
    rowData(17) = SafeNum(txtIV.Value)
    rowData(18) = SafeNum(txtDelta.Value)
    rowData(19) = SafeNum(txtBeta.Value)
    rowData(20) = cmbBroker.Value
    rowData(21) = Trim(txtUnderlying.Value)   ' col U: Underlying (e.g. TWII, SPX)
    rowData(22) = SafeNum(txtGamma.Value)     ' col V: Gamma  (manual entry; RefreshGreeks may overwrite)
    rowData(23) = SafeNum(txtTheta.Value)     ' col W: Theta  (manual entry; RefreshGreeks may overwrite)
    ' 2026-10-18: col X Conv_Ratio was warrant-only; Warrant removed (see
    ' CLAUDE.md). Column kept in the sheet schema (reindexing every generic
    ' column read elsewhere was judged not worth the risk), just never
    ' written again - always blank from here on.
    rowData(24) = ""

    Call RR5_Core.AppendTransaction(rowData)
    MsgBox "Transaction saved: " & rowData(1), vbInformation, "RR5 Transaction"
    Call ClearAll
End Sub

Private Sub btnClear_Click()
    ClearAll
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

' --- Helpers: safe numeric/date conversion ---
Private Function SafeNum(v As Variant) As Variant
    If Trim(CStr(v)) = "" Then
        SafeNum = ""
    Else
        SafeNum = val(v)
    End If
End Function

Private Function SafeDate(v As Variant) As Variant
    On Error Resume Next
    Dim d As Date
    If Trim(CStr(v)) = "" Then
        SafeDate = ""
    Else
        d = CDate(v)
        If Err.Number = 0 Then SafeDate = d Else SafeDate = v
    End If
    On Error GoTo 0
End Function

Private Sub ClearAll()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        ' Skip the rate field -- it is a persistent global setting, not a per-tx field
        If ctrl.Name = "txtRiskFreeRate" Then GoTo NextCtrl
        Select Case TypeName(ctrl)
            Case "TextBox":  ctrl.Value = ""
            Case "ComboBox": ctrl.Value = ""
        End Select
NextCtrl:
    Next ctrl
    txtDate.Value = Format(Date, "yyyy/mm/dd")
    txtFee.Value = "0"
    txtTax.Value = "0"
    txtMultiplier.Value = "1"
    On Error Resume Next
    txtTicker.SetFocus
End Sub


