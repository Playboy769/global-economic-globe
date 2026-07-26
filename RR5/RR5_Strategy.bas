Attribute VB_Name = "RR5_Strategy"
Option Explicit

' =====================================================================
' STRATEGY BUILDER -- Multi-Leg Payoff Calculator (max 6 legs)
' ---------------------------------------------------------------------
' Worksheet: StrategyBuilder
'
' Layout (cols A..K):
'   Row 1   Title
'   Row 3   A "Underlying Symbol:" | B (input symbol e.g. TWII)
'           C "Underlying Price (ref):" | D (input spot)
'           E "Risk-Free Rate:" | F (input, e.g. 0.02)
'           G "Multiplier (default):" | H (input, e.g. 50)
'   Row 4   A "DTE (days):" | B (input, e.g. 30)
'           C "IV (default):" | D (input, e.g. 0.25)
'   Row 5   Note
'   Row 7   Header: Leg# / TYPE / STRIKE / CONTRACTS / MULTIPLIER /
'                   DIRECTION / PREMIUM / DELTA / GAMMA / THETA / VEGA
'   Row 8-13   Leg rows (6 legs)
'   Row 15  "STRATEGY AGGREGATES"
'   Row 16  Net Delta
'   Row 17  Net Gamma
'   Row 18  Net Theta/Day
'   Row 19  Net Vega
'   Row 20  Net Premium Paid     (positive = debit, negative = credit)
'   Row 21  Max Profit (est., at expiry)
'   Row 22  Max Loss (est., at expiry)
'   Row 23  Breakeven (est., at expiry)
'
' Direction sign:
'   LONG  -> +1 (paid premium, you bought the leg)
'   SHORT -> -1 (received premium, you sold the leg)
'
' Aggregation formula:
'   net_Greek  = sum( dirSign * Contracts * Multiplier * leg_Greek )
'   net_Premium = sum( dirSign * Contracts * Multiplier * Premium )
'                 positive = net debit (you paid)
'                 negative = net credit (you received)
'
' Payoff at expiry (per leg):
'   intrinsic = max(0, S - K) for Call,  max(0, K - S) for Put
'   leg_PnL   = dirSign * Contracts * Multiplier * (intrinsic - Premium)
'
' Max Profit / Max Loss / Breakeven detected by scanning S over
' [minStrike * 0.7, maxStrike * 1.3] in 400 steps.
' Unlimited cases flagged when payoff trends past scan boundary.
' =====================================================================

' ---------------------------------------------------------------------
' One-shot layout setup. Run after first install or whenever you want
' to reset labels / defaults. Doesn't touch user inputs that already exist.
' ---------------------------------------------------------------------
Public Sub Setup_StrategyBuilder()
    On Error GoTo Fail
    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Sheets("StrategyBuilder")

    ' --- Row 3 header labels ---
    wsS.Range("A3").Value = "Underlying Symbol:"
    wsS.Range("C3").Value = "Underlying Price (ref):"
    wsS.Range("E3").Value = "Risk-Free Rate:"
    If Not IsNumeric(wsS.Range("F3").Value) Then wsS.Range("F3").Value = 0.02
    wsS.Range("G3").Value = "Multiplier (default):"
    If Not IsNumeric(wsS.Range("H3").Value) Then wsS.Range("H3").Value = 50

    ' --- Row 4 header labels (NEW) ---
    wsS.Range("A4").Value = "DTE (days):"
    wsS.Range("C4").Value = "IV (default):"

    ' --- Row 5 note ---
    wsS.Range("A5").Value = _
        "NOTE: Greeks (H/I/J/K) auto-filled by RefreshStrategyBuilder. " & _
        "Direction = Long or Short. Premium per single share/index point."

    ' --- Row 7 leg table header ---
    Dim hdr As Variant
    hdr = Array("Leg #", "TYPE", "STRIKE", "CONTRACTS", "MULTIPLIER", _
                "DIRECTION", "PREMIUM", "DELTA", "GAMMA", "THETA", "VEGA")
    Dim c As Long
    For c = 0 To UBound(hdr)
        wsS.Cells(7, c + 1).Value = hdr(c)
    Next c

    ' --- Pre-fill leg #s and default multiplier ---
    Dim row As Long
    For row = 8 To 13
        wsS.Cells(row, 1).Value = row - 7
        If Not IsNumeric(wsS.Cells(row, 5).Value) Or wsS.Cells(row, 5).Value = 0 Then
            wsS.Cells(row, 5).Value = wsS.Range("H3").Value
        End If
    Next row

    ' --- Row 15-23 aggregate labels ---
    wsS.Range("A15").Value = "STRATEGY AGGREGATES"
    wsS.Range("A16").Value = "Net Delta"
    wsS.Range("A17").Value = "Net Gamma"
    wsS.Range("A18").Value = "Net Theta/Day"
    wsS.Range("A19").Value = "Net Vega"
    wsS.Range("A20").Value = "Net Premium Paid"
    wsS.Range("A21").Value = "Max Profit (est.)"
    wsS.Range("A22").Value = "Max Loss (est.)"
    wsS.Range("A23").Value = "Breakeven (est.)"

    ' Clear stale results
    wsS.Range("B16:B23").ClearContents

    MsgBox "StrategyBuilder layout ready." & vbCrLf & vbCrLf & _
           "Fill in inputs:" & vbCrLf & _
           "  D3 = Spot price        F3 = Risk-free rate  H3 = Default multiplier" & vbCrLf & _
           "  B4 = DTE (days)        D4 = IV (e.g. 0.25)" & vbCrLf & _
           "  Rows 8-13: TYPE (Call/Put), STRIKE, CONTRACTS, DIRECTION (Long/Short), PREMIUM" & vbCrLf & vbCrLf & _
           "Then Alt+F8 -> RefreshStrategyBuilder.", _
           vbInformation, "StrategyBuilder Setup"
    Exit Sub
Fail:
    MsgBox "Setup_StrategyBuilder failed: " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------
' Main recalculation routine.
' Reads inputs, fills per-leg Greeks (H/I/J/K), writes aggregates (B16-B23).
' ---------------------------------------------------------------------
Public Sub RefreshStrategyBuilder()
    On Error GoTo Fail
    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Sheets("StrategyBuilder")

    ' --- Header inputs ---
    Dim S As Double, rRate As Double, multDef As Double
    Dim dteDays As Double, ivDef As Double, T As Double

    S = ToDbl(wsS.Range("D3").Value)
    rRate = ToDbl(wsS.Range("F3").Value)
    multDef = ToDbl(wsS.Range("H3").Value)
    dteDays = ToDbl(wsS.Range("B4").Value)
    ivDef = ToDbl(wsS.Range("D4").Value)

    If S <= 0 Or dteDays <= 0 Or ivDef <= 0 Then
        MsgBox "Missing required inputs:" & vbCrLf & _
               "  D3 Spot = " & S & vbCrLf & _
               "  B4 DTE  = " & dteDays & vbCrLf & _
               "  D4 IV   = " & ivDef, _
               vbExclamation, "StrategyBuilder"
        Exit Sub
    End If

    T = dteDays / 365#
    If multDef = 0 Then multDef = 50

    ' --- Walk 6 legs ---
    Dim netDelta As Double, netGamma As Double, netTheta As Double
    Dim netVega As Double, netPrem As Double
    Dim n As Long: n = 0

    ' Parallel arrays for payoff scan
    Dim lgType(1 To 6) As String
    Dim lgStrike(1 To 6) As Double
    Dim lgQty(1 To 6) As Double      ' signed weight = dirSign * contracts * mult
    Dim lgPrem(1 To 6) As Double
    Dim lgMinK As Double: lgMinK = 1E+15
    Dim lgMaxK As Double: lgMaxK = 0

    ' Clear stale Greeks
    wsS.Range("H8:K13").ClearContents

    Dim row As Long
    For row = 8 To 13
        Dim tStr As String, kStr As Double, qty As Double, mult As Double
        Dim dirStr As String, prem As Double, dirSign As Long

        tStr = UCase(Trim(CStr(wsS.Cells(row, 2).Value)))
        If tStr = "" Then GoTo NextLeg
        If tStr <> "CALL" And tStr <> "PUT" Then GoTo NextLeg

        kStr = ToDbl(wsS.Cells(row, 3).Value)
        qty = ToDbl(wsS.Cells(row, 4).Value)
        mult = ToDbl(wsS.Cells(row, 5).Value)
        If mult = 0 Then mult = multDef
        dirStr = UCase(Trim(CStr(wsS.Cells(row, 6).Value)))
        prem = ToDbl(wsS.Cells(row, 7).Value)

        If kStr <= 0 Or qty <= 0 Then GoTo NextLeg

        dirSign = 0
        If dirStr = "LONG" Then dirSign = 1
        If dirStr = "SHORT" Then dirSign = -1
        If dirSign = 0 Then GoTo NextLeg   ' direction required

        Dim oType As String
        If tStr = "CALL" Then oType = "Call" Else oType = "Put"

        ' --- BS Greeks per single contract ---
        Dim legDelta As Double, legGamma As Double, legTheta As Double, legVega As Double
        legDelta = BS_Delta(S, kStr, T, rRate, ivDef, oType)
        legGamma = BS_Gamma(S, kStr, T, rRate, ivDef)
        legTheta = BS_Theta(S, kStr, T, rRate, ivDef, oType)
        legVega = BS_Vega(S, kStr, T, rRate, ivDef)

        wsS.Cells(row, 8).Value = legDelta
        wsS.Cells(row, 9).Value = legGamma
        wsS.Cells(row, 10).Value = legTheta
        wsS.Cells(row, 11).Value = legVega
        wsS.Cells(row, 8).NumberFormat = "0.0000"
        wsS.Cells(row, 9).NumberFormat = "0.000000"
        wsS.Cells(row, 10).NumberFormat = "0.0000"
        wsS.Cells(row, 11).NumberFormat = "0.0000"

        ' --- Aggregate (signed × qty × multiplier × per-share Greek) ---
        Dim w As Double: w = dirSign * qty * mult
        netDelta = netDelta + w * legDelta
        netGamma = netGamma + w * legGamma
        netTheta = netTheta + w * legTheta
        netVega = netVega + w * legVega
        netPrem = netPrem + w * prem

        ' --- Cache for payoff scan ---
        n = n + 1
        lgType(n) = oType
        lgStrike(n) = kStr
        lgQty(n) = w
        lgPrem(n) = prem
        If kStr < lgMinK Then lgMinK = kStr
        If kStr > lgMaxK Then lgMaxK = kStr
NextLeg:
    Next row

    ' --- Write aggregates ---
    wsS.Range("B16").Value = netDelta: wsS.Range("B16").NumberFormat = "#,##0.0000"
    wsS.Range("B17").Value = netGamma: wsS.Range("B17").NumberFormat = "#,##0.000000"
    wsS.Range("B18").Value = netTheta: wsS.Range("B18").NumberFormat = "#,##0.00"
    wsS.Range("B19").Value = netVega:  wsS.Range("B19").NumberFormat = "#,##0.00"
    wsS.Range("B20").Value = netPrem:  wsS.Range("B20").NumberFormat = "#,##0;(#,##0)"

    ' --- Payoff scan ---
    If n = 0 Then
        wsS.Range("B21:B23").ClearContents
        Application.StatusBar = "StrategyBuilder: no valid legs"
        Exit Sub
    End If

    Dim scanLo As Double, scanHi As Double
    scanLo = lgMinK * 0.7
    scanHi = lgMaxK * 1.3
    If scanLo < 0 Then scanLo = 0
    Dim stepSize As Double: stepSize = (scanHi - scanLo) / 400
    If stepSize <= 0 Then stepSize = 1

    Dim maxP As Double: maxP = -1E+15
    Dim minP As Double: minP = 1E+15
    Dim maxAtBoundary As Boolean, minAtBoundary As Boolean
    Dim breakevenList As String: breakevenList = ""

    Dim prevPayoff As Double, prevPx As Double
    Dim firstPx As Boolean: firstPx = True

    Dim px As Double: px = scanLo
    Do While px <= scanHi
        Dim payoff As Double
        payoff = PayoffAtExpiry(px, lgType, lgStrike, lgQty, lgPrem, n)

        If payoff > maxP Then
            maxP = payoff
            maxAtBoundary = (px <= scanLo + stepSize) Or (px >= scanHi - stepSize)
        End If
        If payoff < minP Then
            minP = payoff
            minAtBoundary = (px <= scanLo + stepSize) Or (px >= scanHi - stepSize)
        End If

        If Not firstPx Then
            If (prevPayoff < 0 And payoff >= 0) Or (prevPayoff > 0 And payoff <= 0) Then
                If payoff - prevPayoff <> 0 Then
                    Dim be As Double
                    be = prevPx - prevPayoff * (px - prevPx) / (payoff - prevPayoff)
                    If breakevenList = "" Then
                        breakevenList = Format(be, "#,##0.00")
                    Else
                        breakevenList = breakevenList & " / " & Format(be, "#,##0.00")
                    End If
                End If
            End If
        End If

        prevPayoff = payoff
        prevPx = px
        firstPx = False
        px = px + stepSize
    Loop

    ' --- Unlimited detection (gradient at boundaries) ---
    Dim payLo As Double, payLoPlus As Double
    Dim payHi As Double, payHiMinus As Double
    payLo = PayoffAtExpiry(scanLo, lgType, lgStrike, lgQty, lgPrem, n)
    payLoPlus = PayoffAtExpiry(scanLo + stepSize, lgType, lgStrike, lgQty, lgPrem, n)
    payHiMinus = PayoffAtExpiry(scanHi - stepSize, lgType, lgStrike, lgQty, lgPrem, n)
    payHi = PayoffAtExpiry(scanHi, lgType, lgStrike, lgQty, lgPrem, n)

    Dim gradHi As Double: gradHi = payHi - payHiMinus
    Dim gradLo As Double: gradLo = payLoPlus - payLo

    Dim maxStr As String, minStr As String
    If maxAtBoundary And (gradHi > 0.01 Or gradLo < -0.01) Then
        maxStr = "UNLIMITED (scan caps ~" & Format(maxP, "#,##0") & ")"
    Else
        maxStr = Format(maxP, "#,##0")
    End If

    If minAtBoundary And (gradHi < -0.01 Or gradLo > 0.01) Then
        minStr = "UNLIMITED LOSS (scan caps ~" & Format(minP, "#,##0") & ")"
    Else
        minStr = Format(minP, "#,##0")
    End If

    wsS.Range("B21").Value = maxStr
    wsS.Range("B22").Value = minStr
    wsS.Range("B23").Value = IIf(breakevenList = "", "None in scan range", breakevenList)

    ' F2: redraw payoff chart with the same legs
    Call BuildPayoffChart

    Application.StatusBar = "StrategyBuilder refreshed " & Format(Now, "hh:mm:ss") & _
        "  |  Legs=" & n & "  |  Scan " & Format(scanLo, "0") & "~" & Format(scanHi, "0")
    Exit Sub
Fail:
    Application.StatusBar = False
    MsgBox "RefreshStrategyBuilder failed: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' F2: PAYOFF CHART
' ---------------------------------------------------------------------
' Builds a dynamic P&L-at-expiry curve in the StrategyBuilder sheet.
' Writes data table at A30:B82, then creates/updates an XY scatter chart
' anchored at cell M2 (~col M, top row).
'
' Stage-by-stage error reporting -- on failure MsgBox shows which stage.
' =====================================================================
Public Sub BuildPayoffChart()
    Dim stage As String: stage = "init"
    On Error GoTo Fail

    stage = "open sheet"
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("StrategyBuilder")

    stage = "read defaults"
    Dim multDef As Double
    multDef = ToDbl(wsS.Range("H3").Value)
    If multDef = 0 Then multDef = 50

    stage = "collect legs"
    Dim nL As Long: nL = 0
    Dim lgType(1 To 6) As String
    Dim lgStrike(1 To 6) As Double
    Dim lgQty(1 To 6) As Double
    Dim lgPrem(1 To 6) As Double
    Dim minK As Double: minK = 1E+15
    Dim maxK As Double: maxK = 0

    Dim r As Long
    For r = 8 To 13
        stage = "leg row " & r
        Dim tStr As String: tStr = UCase(Trim(CStr(wsS.Cells(r, 2).Value)))
        If tStr <> "CALL" And tStr <> "PUT" Then GoTo NextLeg
        Dim k As Double: k = ToDbl(wsS.Cells(r, 3).Value)
        Dim qq As Double: qq = ToDbl(wsS.Cells(r, 4).Value)
        Dim mm As Double: mm = ToDbl(wsS.Cells(r, 5).Value)
        If mm = 0 Then mm = multDef
        Dim d As String: d = UCase(Trim(CStr(wsS.Cells(r, 6).Value)))
        Dim pr As Double: pr = ToDbl(wsS.Cells(r, 7).Value)
        If k <= 0 Or qq <= 0 Then GoTo NextLeg
        ' dirSign instead of sgn to avoid clash with VBA built-in Sgn()
        Dim dirSign As Long
        If d = "LONG" Then
            dirSign = 1
        ElseIf d = "SHORT" Then
            dirSign = -1
        Else
            GoTo NextLeg
        End If
        nL = nL + 1
        If tStr = "CALL" Then lgType(nL) = "Call" Else lgType(nL) = "Put"
        lgStrike(nL) = k
        lgQty(nL) = dirSign * qq * mm
        lgPrem(nL) = pr
        If k < minK Then minK = k
        If k > maxK Then maxK = k
NextLeg:
    Next r

    stage = "validate legs"
    wsS.Range("A30:B82").ClearContents
    wsS.Cells(30, 1).Value = "PAYOFF DATA"
    If nL = 0 Then
        ' Remove existing chart if no legs
        On Error Resume Next
        wsS.ChartObjects("PayoffChart").Delete
        On Error GoTo Fail
        Exit Sub
    End If

    stage = "compute scan range"
    Dim loSpot As Double, hiSpot As Double, stepSz As Double
    loSpot = minK * 0.7
    hiSpot = maxK * 1.3
    If loSpot < 0 Then loSpot = 0
    If hiSpot <= loSpot Then hiSpot = loSpot + 100
    stepSz = (hiSpot - loSpot) / 50
    If stepSz <= 0 Then stepSz = 1

    stage = "write payoff data"
    wsS.Cells(31, 1).Value = "Spot"
    wsS.Cells(31, 2).Value = "PnL"
    Dim sp As Double: sp = loSpot
    Dim outR As Long: outR = 32
    Do While sp <= hiSpot + 0.0001
        Dim payoff As Double: payoff = 0
        Dim i As Long
        For i = 1 To nL
            Dim intr As Double
            If lgType(i) = "Call" Then
                intr = sp - lgStrike(i)
                If intr < 0 Then intr = 0
            Else
                intr = lgStrike(i) - sp
                If intr < 0 Then intr = 0
            End If
            payoff = payoff + lgQty(i) * (intr - lgPrem(i))
        Next i
        wsS.Cells(outR, 1).Value = sp
        wsS.Cells(outR, 2).Value = payoff
        outR = outR + 1
        sp = sp + stepSz
    Loop
    Dim lastR As Long: lastR = outR - 1
    If lastR < 33 Then Exit Sub

    stage = "find/create chart"
    Dim co As ChartObject
    Dim foundChart As Boolean: foundChart = False
    Dim cc As ChartObject
    For Each cc In wsS.ChartObjects
        If cc.Name = "PayoffChart" Then
            Set co = cc
            foundChart = True
            Exit For
        End If
    Next cc
    If Not foundChart Then
        Dim anchorL As Double, anchorT As Double
        anchorL = wsS.Range("M2").Left
        anchorT = wsS.Range("M2").Top
        Set co = wsS.ChartObjects.Add(Left:=anchorL, Top:=anchorT, _
                                       Width:=460, Height:=300)
        co.Name = "PayoffChart"
    End If

    stage = "configure chart"
    With co.Chart
        .ChartType = xlXYScatterLines
        ' Clear existing series
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        ' Add PnL series
        Dim sr As Series: Set sr = .SeriesCollection.NewSeries
        sr.XValues = wsS.Range(wsS.Cells(32, 1), wsS.Cells(lastR, 1))
        sr.Values = wsS.Range(wsS.Cells(32, 2), wsS.Cells(lastR, 2))
        sr.Name = "P&L at Expiry"
        ' Zero line
        Dim sr0 As Series: Set sr0 = .SeriesCollection.NewSeries
        sr0.XValues = Array(loSpot, hiSpot)
        sr0.Values = Array(0, 0)
        sr0.Name = "Zero"
        On Error Resume Next
        sr0.Format.Line.ForeColor.RGB = RGB(150, 150, 150)
        sr0.Format.Line.DashStyle = msoLineDash
        On Error GoTo Fail

        .HasTitle = True
        .ChartTitle.Text = "Payoff at Expiry  (Spot " & _
            Format(loSpot, "0") & " ~ " & Format(hiSpot, "0") & ")"
        .HasLegend = False
        On Error Resume Next
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Spot Price"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "P&L"
        On Error GoTo Fail
    End With

    Exit Sub
Fail:
    MsgBox "BuildPayoffChart failed at stage [" & stage & "]:" & vbCrLf & _
           Err.Description, vbExclamation, "F2 Payoff Chart"
End Sub

' =====================================================================
' F3: STRATEGY TEMPLATES -- one-click leg pre-fill
' ---------------------------------------------------------------------
' Each template reads D3 (Spot) + H3 (Default Multiplier) + an optional
' "width" parameter (distance between strikes) and writes legs 1-N
' into rows 8-13. Calls RefreshStrategyBuilder at the end.
'
' Strike rounding:
'   TWII / TXO  -> nearest 100 (typical TXO strike grid)
'   2330 / Equity -> nearest 50 or 10 depending on spot
'   Otherwise   -> nearest 1
'
' All templates default to qty = 1 contract, premium = 0 (user fills
' actual premiums after pre-fill).
' =====================================================================

' --- Helper: round strike to product grid ---
Private Function RoundStrike(spot As Double) As Double
    Dim grid As Double
    Select Case True
        Case spot >= 10000:  grid = 100   ' TXO / TWII
        Case spot >= 1000:   grid = 50
        Case spot >= 100:    grid = 10
        Case Else:           grid = 1
    End Select
    RoundStrike = Application.WorksheetFunction.Round(spot / grid, 0) * grid
End Function

Private Function StrikeGrid(spot As Double) As Double
    Select Case True
        Case spot >= 10000:  StrikeGrid = 100
        Case spot >= 1000:   StrikeGrid = 50
        Case spot >= 100:    StrikeGrid = 10
        Case Else:           StrikeGrid = 1
    End Select
End Function

' --- Helper: clear all leg rows ---
Private Sub ClearLegRows(wsS As Worksheet)
    Dim r As Long
    For r = 8 To 13
        wsS.Range(wsS.Cells(r, 2), wsS.Cells(r, 11)).ClearContents
        wsS.Cells(r, 1).Value = r - 7   ' restore leg #
    Next r
End Sub

' --- Helper: write one leg ---
Private Sub WriteLeg(wsS As Worksheet, legRow As Long, _
                     legType As String, strike As Double, _
                     qty As Long, direction As String, premium As Double, _
                     Optional mult As Variant)
    wsS.Cells(legRow, 2).Value = legType
    wsS.Cells(legRow, 3).Value = strike
    wsS.Cells(legRow, 4).Value = qty
    If Not IsMissing(mult) Then
        wsS.Cells(legRow, 5).Value = mult
    Else
        wsS.Cells(legRow, 5).Value = ToDbl(wsS.Range("H3").Value)
    End If
    wsS.Cells(legRow, 6).Value = direction
    wsS.Cells(legRow, 7).Value = premium
End Sub

' --- Template helper: validate spot ---
Private Function GetSpotOrFail(wsS As Worksheet, label As String) As Double
    GetSpotOrFail = ToDbl(wsS.Range("D3").Value)
    If GetSpotOrFail <= 0 Then
        MsgBox "Set D3 (Underlying Price) before using template [" & label & "].", _
               vbExclamation, "Strategy Template"
    End If
End Function

' =====================================================================
' TEMPLATE 1: Bull Call Spread
'   Long  Call @ ATM
'   Short Call @ ATM + 1 grid
' Directional bullish, limited risk + limited reward
' =====================================================================
Public Sub Tmpl_BullCallSpread()
    Dim stage As String: stage = "init"
    On Error GoTo Fail
    stage = "open sheet"
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("StrategyBuilder")
    stage = "get spot"
    Dim S As Double: S = GetSpotOrFail(wsS, "Bull Call Spread")
    If S = 0 Then Exit Sub
    stage = "compute strikes"
    Dim atm As Double: atm = RoundStrike(S)
    Dim g As Double: g = StrikeGrid(S)
    stage = "write legs"
    Call ClearLegRows(wsS)
    Call WriteLeg(wsS, 8, "Call", atm, 1, "Long", 0)
    Call WriteLeg(wsS, 9, "Call", atm + g, 1, "Short", 0)
    stage = "refresh"
    Call RefreshStrategyBuilder
    Exit Sub
Fail:
    MsgBox "Tmpl_BullCallSpread failed at [" & stage & "]: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' TEMPLATE 2: Bear Put Spread
'   Long  Put @ ATM
'   Short Put @ ATM - 1 grid
' =====================================================================
Public Sub Tmpl_BearPutSpread()
    Dim stage As String: stage = "init"
    On Error GoTo Fail
    stage = "open sheet"
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("StrategyBuilder")
    stage = "get spot"
    Dim S As Double: S = GetSpotOrFail(wsS, "Bear Put Spread")
    If S = 0 Then Exit Sub
    stage = "compute strikes"
    Dim atm As Double: atm = RoundStrike(S)
    Dim g As Double: g = StrikeGrid(S)
    stage = "write legs"
    Call ClearLegRows(wsS)
    Call WriteLeg(wsS, 8, "Put", atm, 1, "Long", 0)
    Call WriteLeg(wsS, 9, "Put", atm - g, 1, "Short", 0)
    stage = "refresh"
    Call RefreshStrategyBuilder
    Exit Sub
Fail:
    MsgBox "Tmpl_BearPutSpread failed at [" & stage & "]: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' TEMPLATE 3: Iron Condor (1 grid wings, 1 grid inner)
'   Short Put  @ ATM - 1g
'   Long  Put  @ ATM - 2g
'   Short Call @ ATM + 1g
'   Long  Call @ ATM + 2g
' Range-bound, sell premium, defined risk
' =====================================================================
Public Sub Tmpl_IronCondor()
    Dim stage As String: stage = "init"
    On Error GoTo Fail
    stage = "open sheet"
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("StrategyBuilder")
    stage = "get spot"
    Dim S As Double: S = GetSpotOrFail(wsS, "Iron Condor")
    If S = 0 Then Exit Sub
    stage = "compute strikes"
    Dim atm As Double: atm = RoundStrike(S)
    Dim g As Double: g = StrikeGrid(S)
    stage = "write legs"
    Call ClearLegRows(wsS)
    Call WriteLeg(wsS, 8, "Put",  atm - g,     1, "Short", 0)
    Call WriteLeg(wsS, 9, "Put",  atm - 2 * g, 1, "Long",  0)
    Call WriteLeg(wsS, 10, "Call", atm + g,     1, "Short", 0)
    Call WriteLeg(wsS, 11, "Call", atm + 2 * g, 1, "Long",  0)
    stage = "refresh"
    Call RefreshStrategyBuilder
    Exit Sub
Fail:
    MsgBox "Tmpl_IronCondor failed at [" & stage & "]: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' TEMPLATE 4: Long Straddle
'   Long Call @ ATM
'   Long Put  @ ATM
' Volatility play, unlimited upside, defined downside (= premium paid)
' =====================================================================
Public Sub Tmpl_LongStraddle()
    Dim stage As String: stage = "init"
    On Error GoTo Fail
    stage = "open sheet"
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("StrategyBuilder")
    stage = "get spot"
    Dim S As Double: S = GetSpotOrFail(wsS, "Long Straddle")
    If S = 0 Then Exit Sub
    stage = "compute strike"
    Dim atm As Double: atm = RoundStrike(S)
    stage = "write legs"
    Call ClearLegRows(wsS)
    Call WriteLeg(wsS, 8, "Call", atm, 1, "Long", 0)
    Call WriteLeg(wsS, 9, "Put",  atm, 1, "Long", 0)
    stage = "refresh"
    Call RefreshStrategyBuilder
    Exit Sub
Fail:
    MsgBox "Tmpl_LongStraddle failed at [" & stage & "]: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' TEMPLATE 5: Short Strangle
'   Short Call @ ATM + 1g
'   Short Put  @ ATM - 1g
' Range-bound, sell premium, UNLIMITED risk (be careful)
' =====================================================================
Public Sub Tmpl_ShortStrangle()
    Dim stage As String: stage = "init"
    On Error GoTo Fail
    stage = "open sheet"
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("StrategyBuilder")
    stage = "get spot"
    Dim S As Double: S = GetSpotOrFail(wsS, "Short Strangle")
    If S = 0 Then Exit Sub
    stage = "compute strikes"
    Dim atm As Double: atm = RoundStrike(S)
    Dim g As Double: g = StrikeGrid(S)
    stage = "write legs"
    Call ClearLegRows(wsS)
    Call WriteLeg(wsS, 8, "Call", atm + g, 1, "Short", 0)
    Call WriteLeg(wsS, 9, "Put",  atm - g, 1, "Short", 0)
    stage = "refresh"
    Call RefreshStrategyBuilder
    Exit Sub
Fail:
    MsgBox "Tmpl_ShortStrangle failed at [" & stage & "]: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' TEMPLATE 6: Calendar Spread (same strike, different DTE)
'   Short Call @ ATM (near-term -- user adjusts DTE manually)
'   Long  Call @ ATM (longer-term)
' NOTE: This builder uses ONE shared DTE in B4. Calendar spreads
' require per-leg DTE which isn't modeled yet. The output is informational.
' =====================================================================
Public Sub Tmpl_CalendarSpread()
    Dim stage As String: stage = "init"
    On Error GoTo Fail
    stage = "open sheet"
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("StrategyBuilder")
    stage = "get spot"
    Dim S As Double: S = GetSpotOrFail(wsS, "Calendar Spread")
    If S = 0 Then Exit Sub
    stage = "compute strikes"
    Dim atm As Double: atm = RoundStrike(S)
    stage = "write legs"
    Call ClearLegRows(wsS)
    Call WriteLeg(wsS, 8, "Call", atm, 1, "Short", 0)
    Call WriteLeg(wsS, 9, "Call", atm, 1, "Long",  0)
    stage = "refresh"
    Call RefreshStrategyBuilder
    MsgBox "Calendar Spread pre-filled (same strike, both Call)." & vbCrLf & _
           "NOTE: This builder uses ONE shared DTE/IV. Calendar spread" & vbCrLf & _
           "Greeks are approximate -- both legs share the SAME expiry here.", _
           vbInformation, "Calendar Spread"
    Exit Sub
Fail:
    MsgBox "Tmpl_CalendarSpread failed at [" & stage & "]: " & Err.Description, vbExclamation
End Sub

' ---------------------------------------------------------------------
' Compute strategy payoff at a single underlying price (at expiry).
' Inputs are pre-cached arrays from RefreshStrategyBuilder.
' ---------------------------------------------------------------------
Private Function PayoffAtExpiry(px As Double, _
    lgType() As String, lgStrike() As Double, _
    lgQty() As Double, lgPrem() As Double, n As Long) As Double
    Dim total As Double: total = 0
    Dim i As Long
    For i = 1 To n
        Dim intr As Double
        If lgType(i) = "Call" Then
            intr = px - lgStrike(i)
            If intr < 0 Then intr = 0
        Else
            intr = lgStrike(i) - px
            If intr < 0 Then intr = 0
        End If
        ' lgQty already signed: long=+, short=-
        total = total + lgQty(i) * (intr - lgPrem(i))
    Next i
    PayoffAtExpiry = total
End Function

' ---------------------------------------------------------------------
' Safe numeric conversion: blank / text -> 0
' ---------------------------------------------------------------------
Private Function ToDbl(v As Variant) As Double
    If IsNumeric(v) Then
        ToDbl = CDbl(v)
    Else
        ToDbl = 0
    End If
End Function
