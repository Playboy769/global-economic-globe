Attribute VB_Name = "RR5_Dashboard"
Option Explicit

' =====================================================================
' RR5 DASHBOARD AGGREGATIONS
' ---------------------------------------------------------------------
' One-shot setup:
'    FixWorksheetFormulas    -- patches 4 bugs in B1 / D2 / F2 / H2
'
' Main routine:
'    DashboardRefresh        -- pulls Realized / IVTracker / HistoryLog /
'                               Greeks_History aggregates, writes rows 3-4
'
' Helpers (private): WriteRealizedMetrics / WriteIVMetrics /
'                    WriteHistoryLogMetrics / WriteGreeksHistoryMetrics /
'                    WritePositionMetrics / EnsureDashboardLabels
' =====================================================================

' =====================================================================
' ONE-SHOT: Archive GreeksView (RR5 now replaces it)
' ---------------------------------------------------------------------
' Renames "GreeksView" -> "_GreeksView_Archive" and hides it via
' xlSheetVeryHidden (only re-shown by editing VBA properties).
' Non-destructive: undo by running RestoreGreeksView.
' =====================================================================
Public Sub ArchiveGreeksView()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("GreeksView")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "GreeksView already archived or never existed.", vbInformation
        Exit Sub
    End If

    On Error GoTo Fail
    ws.Name = "_GreeksView_Archive"
    ws.Visible = xlSheetVeryHidden    ' xlSheetVeryHidden = 2

    MsgBox "GreeksView archived as '_GreeksView_Archive' (very hidden)." & vbCrLf & _
           "Run RestoreGreeksView to undo.", vbInformation
    Exit Sub
Fail:
    MsgBox "ArchiveGreeksView failed: " & Err.Description, vbCritical
End Sub

Public Sub RestoreGreeksView()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("_GreeksView_Archive")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No archived GreeksView found.", vbInformation
        Exit Sub
    End If

    On Error GoTo Fail
    ws.Visible = xlSheetVisible
    ws.Name = "GreeksView"
    MsgBox "GreeksView restored.", vbInformation
    Exit Sub
Fail:
    MsgBox "RestoreGreeksView failed: " & Err.Description, vbCritical
End Sub

' =====================================================================
' ONE-SHOT: Rewrite summary-bar formulas to COMPACT layout (8 rows/sec)
' ---------------------------------------------------------------------
'   Futures  data rows 7-14   (subtotal 15)
'   Options  data rows 19-26  (subtotal 27)
'   Warrants data rows 31-38  (subtotal 39)
'   GrandTotal row 41
' =====================================================================
Public Sub FixWorksheetFormulas()
    On Error GoTo Fail
    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Sheets("RR5")

    ' --- Row 1 (label / Total MV / Net Delta / Net Gamma / Positions) ---
    wsR.Range("B1").Formula = _
        "=IFERROR(SUBTOTAL(9,N7:N14)+SUBTOTAL(9,Q19:Q26)+SUBTOTAL(9,Q31:Q38),0)"
    wsR.Range("L1").Formula = _
        "=IFERROR(SUBTOTAL(9,S7:S14)+SUBTOTAL(9,Y19:Y26)" & _
        "+SUMPRODUCT((A31:A38<>"""")*M31:M38*T31:T38),0)"
    wsR.Range("N1").Formula = _
        "=IFERROR(SUMPRODUCT((H19:H26<>"""")*U19:U26*L19:L26*M19:M26)+X39,0)"
    wsR.Range("P1").Formula = "=COUNTA(A7:A14)+COUNTA(A19:A26)+COUNTA(A31:A38)"

    ' --- Row 2 (UNRL PNL% / UNRL PNL / RL PNL / NET THETA / NET VEGA / MARGIN) ---
    wsR.Range("D2").Formula = _
        "=IFERROR((SUBTOTAL(9,O7:O14)+SUBTOTAL(9,R19:R26)+SUBTOTAL(9,R31:R38))/B1,0)"
    wsR.Range("F2").Formula = _
        "=IFERROR(SUBTOTAL(9,O7:O14)+SUBTOTAL(9,R19:R26)+SUBTOTAL(9,R31:R38),0)"
    wsR.Range("H2").Formula = "=IFERROR(SUM(Realized!I:I),0)"
    wsR.Range("J2").Formula = "=IFERROR(SUBTOTAL(9,Z19:Z26)+Y39,0)"
    wsR.Range("L2").Formula = "=IFERROR(SUBTOTAL(9,AA19:AA26)+Z39,0)"
    wsR.Range("N2").Formula = "=IFERROR(SUBTOTAL(9,Q7:Q14),0)"

    MsgBox "Summary bar formulas rewritten for compact layout:" & vbCrLf & _
           "  Futures  N7:N14  (NET EXPOS), O7:O14 (UNRL PNL)" & vbCrLf & _
           "  Options  Q19:Q26 (NET EXPOS), R19:R26 (UNRL PNL)" & vbCrLf & _
           "  Warrants Q31:Q38 (NET EXPOS), R31:R38 (UNRL PNL)" & vbCrLf & vbCrLf & _
           "Don't forget to also update:" & vbCrLf & _
           "  - row 15 Subtotal Futures   formulas" & vbCrLf & _
           "  - row 27 Subtotal Options   formulas" & vbCrLf & _
           "  - row 39 Subtotal Warrants  formulas (X/Y/Z need SUMPRODUCT)" & vbCrLf & _
           "  - row 41 Grand Total        formulas", _
           vbInformation, "Worksheet Patch OK"
    Exit Sub
Fail:
    MsgBox "FixWorksheetFormulas failed: " & Err.Description, vbCritical
End Sub

' =====================================================================
' ONE-SHOT: Write all PER-ROW formulas at compact layout rows.
' ---------------------------------------------------------------------
' Restores DTE / % CHG / NET EXPOS / UNRL PNL / NET DELTA / WT% / etc.
' formulas at the new positions (Fut 7-14, Opt 19-26, Wrt 31-38).
' Excel auto-adjusts row references when Formula is set on a range,
' so each cell ends up with the correct row number.
' =====================================================================
Public Sub Setup_CompactRowFormulas()
    On Error GoTo Fail
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets("RR5")

    ' ---------- Futures rows 7-14 ----------
    wsR.Range("E7:E14").Formula = _
        "=IF(D7="""","""",MAX(0,D7-TODAY()))"
    wsR.Range("M7:M14").Formula = _
        "=IF(OR(K7="""",L7=""""),"""",IFERROR((L7-K7)/K7,""""))"
    wsR.Range("N7:N14").Formula = _
        "=IF(OR(I7="""",J7="""",L7=""""),"""",IF(H7=""Long"",I7*J7*L7,-I7*J7*L7))"
    wsR.Range("O7:O14").Formula = _
        "=IF(OR(I7="""",J7="""",K7="""",L7=""""),"""",IF(H7=""Long"",I7*J7*(L7-K7),-I7*J7*(L7-K7)))"
    wsR.Range("Q7:Q14").Formula = _
        "=IF(OR(I7="""",P7=""""),"""",I7*P7)"
    wsR.Range("S7:S14").Formula = _
        "=IF(OR(I7="""",J7="""",R7=""""),"""",IF(H7=""Long"",I7*J7*R7,-I7*J7*R7))"
    wsR.Range("V7:V14").Formula = _
        "=IF(OR(N7="""",'RR5'!$B$1=0),"""",N7/'RR5'!$B$1)"

    ' ---------- Options rows 19-26 ----------
    wsR.Range("E19:E26").Formula = _
        "=IF(D19="""","""",MAX(0,D19-TODAY()))"
    wsR.Range("P19:P26").Formula = _
        "=IF(OR(N19="""",O19=""""),"""",IFERROR((O19-N19)/N19,""""))"
    wsR.Range("Q19:Q26").Formula = _
        "=IF(OR(L19="""",M19="""",O19=""""),""""," & _
        "IF(G19=""CBOE"",L19*M19*O19*'RR5'!$D$1,L19*M19*O19))"
    wsR.Range("R19:R26").Formula = _
        "=IF(OR(L19="""",M19="""",N19="""",O19=""""),""""," & _
        "IF(G19=""CBOE"",L19*M19*(O19-N19)*'RR5'!$D$1,L19*M19*(O19-N19)))"
    wsR.Range("Y19:Y26").Formula = _
        "=IF(OR(L19="""",M19="""",T19=""""),""""," & _
        "IF(H19=""Put"",-L19*M19*ABS(T19),L19*M19*T19))"
    wsR.Range("Z19:Z26").Formula = _
        "=IF(OR(L19="""",M19="""",V19=""""),"""",L19*M19*V19)"
    wsR.Range("AA19:AA26").Formula = _
        "=IF(OR(L19="""",M19="""",W19=""""),"""",L19*M19*W19)"
    wsR.Range("AC19:AC26").Formula = _
        "=IF(OR(Q19="""",'RR5'!$B$1=0),"""",Q19/'RR5'!$B$1)"

    ' ---------- Warrants rows 31-38 ----------
    wsR.Range("E31:E38").Formula = _
        "=IF(D31="""","""",MAX(0,D31-TODAY()))"
    wsR.Range("P31:P38").Formula = _
        "=IF(OR(N31="""",O31=""""),"""",IFERROR((O31-N31)/N31,""""))"
    wsR.Range("Q31:Q38").Formula = _
        "=IF(OR(M31="""",O31=""""),"""",M31*O31)"
    wsR.Range("R31:R38").Formula = _
        "=IF(OR(M31="""",N31="""",O31=""""),"""",M31*(O31-N31))"
    wsR.Range("U31:U38").Formula = _
        "=IF(OR(O31="""",J31="""",K31="""",L31="""",O31=0),""""," & _
        "IFERROR((O31-MAX(0,(J31-K31)/L31))/O31,""""))"
    wsR.Range("V31:V38").Formula = "=E31"
    wsR.Range("W31:W38").Formula = _
        "=IF(OR(Q31="""",'RR5'!$B$1=0),"""",Q31/'RR5'!$B$1)"

    MsgBox "Per-row formulas restored:" & vbCrLf & _
           "  Futures  rows 7-14   (E/M/N/O/Q/S/V)" & vbCrLf & _
           "  Options  rows 19-26  (E/P/Q/R/Y/Z/AA/AC)" & vbCrLf & _
           "  Warrants rows 31-38  (E/P/Q/R/U/V/W)" & vbCrLf & vbCrLf & _
           "Now DTE/% CHG/NET EXPOS/UNRL PNL/WT% etc. auto-compute.", _
           vbInformation, "Compact Row Formulas OK"
    Exit Sub
Fail:
    MsgBox "Setup_CompactRowFormulas failed: " & Err.Description, vbCritical
End Sub

' =====================================================================
' ONE-SHOT: Write subtotal + grand-total formulas at the COMPACT layout
' rows (15, 27, 39, 41). Run AFTER FixWorksheetFormulas.
' =====================================================================
Public Sub Setup_CompactSubtotals()
    On Error GoTo Fail
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets("RR5")

    ' --- Row 15: Futures subtotal (N=NET EXPOS, O=UNRL PNL, Q=MARGIN, S=NET DELTA, V=WT%) ---
    wsR.Range("A15").Value = "  Subtotal  Futures"
    wsR.Range("N15").Formula = "=SUBTOTAL(9,N7:N14)"
    wsR.Range("O15").Formula = "=SUBTOTAL(9,O7:O14)"
    wsR.Range("Q15").Formula = "=SUBTOTAL(9,Q7:Q14)"
    wsR.Range("S15").Formula = "=SUBTOTAL(9,S7:S14)"
    wsR.Range("V15").Formula = "=SUBTOTAL(9,V7:V14)"

    ' --- Row 27: Options subtotal (Q=NET EXPOS, R=UNRL PNL, Y=NET DELTA, Z=NET THETA, AA=NET VEGA, AC=WT%) ---
    wsR.Range("A27").Value = "  Subtotal  Options"
    wsR.Range("Q27").Formula = "=SUBTOTAL(9,Q19:Q26)"
    wsR.Range("R27").Formula = "=SUBTOTAL(9,R19:R26)"
    wsR.Range("Y27").Formula = "=SUBTOTAL(9,Y19:Y26)"
    wsR.Range("Z27").Formula = "=SUBTOTAL(9,Z19:Z26)"
    wsR.Range("AA27").Formula = "=SUBTOTAL(9,AA19:AA26)"
    wsR.Range("AC27").Formula = "=SUBTOTAL(9,AC19:AC26)"

    ' --- Row 39: Warrants subtotal (Q=NET EXPOS, R=UNRL PNL, W=WT%, X/Y/Z=SUMPRODUCT Greeks) ---
    wsR.Range("A39").Value = "  Subtotal  Warrants"
    wsR.Range("Q39").Formula = "=SUBTOTAL(9,Q31:Q38)"
    wsR.Range("R39").Formula = "=SUBTOTAL(9,R31:R38)"
    wsR.Range("W39").Formula = "=SUBTOTAL(9,W31:W38)"
    ' SUMPRODUCT for warrant Greeks (SHARES × per-warrant Greek)
    wsR.Range("X39").Formula = "=SUMPRODUCT((A31:A38<>"""")*M31:M38*X31:X38)"
    wsR.Range("Y39").Formula = "=SUMPRODUCT((A31:A38<>"""")*M31:M38*Y31:Y38)"
    wsR.Range("Z39").Formula = "=SUMPRODUCT((A31:A38<>"""")*M31:M38*Z31:Z38)"

    ' --- Row 41: Grand Total (sum of the three subtotals) ---
    wsR.Range("A41").Value = "  GRAND TOTAL"
    wsR.Range("N41").Formula = "=N15+Q27+Q39"     ' Total NET EXPOS TWD
    wsR.Range("O41").Formula = "=O15+R27+R39"     ' Total UNRL PNL TWD

    MsgBox "Compact subtotals written:" & vbCrLf & _
           "  Row 15: Subtotal Futures"  & vbCrLf & _
           "  Row 27: Subtotal Options"  & vbCrLf & _
           "  Row 39: Subtotal Warrants (+ SUMPRODUCT for X/Y/Z Greeks)" & vbCrLf & _
           "  Row 41: Grand Total" & vbCrLf & vbCrLf & _
           "(You can still style row heights / banners visually as you like.)", _
           vbInformation, "Compact Subtotals OK"
    Exit Sub
Fail:
    MsgBox "Setup_CompactSubtotals failed: " & Err.Description, vbCritical
End Sub

' =====================================================================
' MAIN: DashboardRefresh
' =====================================================================
Public Sub DashboardRefresh()
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Sheets("RR5")

    Call EnsureDashboardLabels(wsR)
    Call WriteRealizedMetrics(wsR)
    Call WriteIVMetrics(wsR)
    Call WriteHistoryLogMetrics(wsR)
    Call WriteGreeksHistoryMetrics(wsR)
    Call WritePositionMetrics(wsR)
    Call RefreshAnalysis           ' Per-position + DTE buckets + Strategy aggregation
    Call BuildScenarioAnalysis     ' F1: spot/IV scenario P&L grid
    Call BuildEquityDrawdownCharts ' F4: equity curve + drawdown

    Application.ScreenUpdating = True
    Application.StatusBar = "DashboardRefresh OK  " & Format(Now, "hh:mm:ss")
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "DashboardRefresh failed: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' RefreshAnalysis -- writes Analysis sheet per-position table + aggregates
' ---------------------------------------------------------------------
' Reads RR5 sections (Futures 7-14 / Options 19-26 / Warrants 31-38),
' writes one row per active position into Analysis starting at row 15
' (table headers are at row 14), then appends derived sections:
'   - DTE Bucket  (rows 40-44):  count + |delta| + theta exposure
'   - Strategy    (rows 46-55):  per strategy tag count + net delta + theta
'   - Underlying  (rows 57-63):  per underlying notional + delta exposure
' =====================================================================
Public Sub RefreshAnalysis()
    On Error GoTo Fail
    Dim wsA As Worksheet, wsR As Worksheet
    Set wsA = ThisWorkbook.Sheets("Analysis")
    Set wsR = ThisWorkbook.Sheets("RR5")

    ' Clear old data values only (do NOT touch user's formatting)
    ' D2:Q13 = new top-area summary sections (DTE / Strategy / Underlying)
    ' A15:Q70 = per-position table + any legacy bottom sections
    wsA.Range("D2:Q13").ClearContents
    wsA.Range("A15:Q70").ClearContents

    Dim r As Long: r = 15

    ' --- Aggregation buckets while we walk positions ---
    Dim dteCnt(0 To 3) As Long, dteAbsDelta(0 To 3) As Double, dteTheta(0 To 3) As Double
    Dim stratDict As Object: Set stratDict = CreateObject("Scripting.Dictionary")
    Dim undlDict As Object: Set undlDict = CreateObject("Scripting.Dictionary")

    ' --- Walk 3 sections ---
    Dim sections As Variant
    sections = Array(Array(7, 14, "fut"), Array(19, 26, "opt"), Array(31, 38, "wrt"))
    Dim sec As Variant, i As Long

    For Each sec In sections
        Dim startR As Long: startR = CLng(sec(0))
        Dim endR As Long:   endR = CLng(sec(1))
        Dim kind As String: kind = CStr(sec(2))

        For i = startR To endR
            Dim code As String
            code = Trim(CStr(wsR.Cells(i, 1).Value))
            If code = "" Then GoTo NextRow

            ' Common fields
            wsA.Cells(r, 1).Value = code
            wsA.Cells(r, 2).Value = wsR.Cells(i, 2).Value

            Dim entryDt As Variant: entryDt = wsR.Cells(i, 3).Value
            If IsDate(entryDt) Then wsA.Cells(r, 4).Value = Date - CDate(entryDt)

            wsA.Cells(r, 8).Value = wsR.Cells(i, 5).Value   ' H DTE

            Dim typStr As String, undl As Variant, undlPx As Variant
            Dim strike As Variant, chgPct As Variant
            Dim delta As Variant, gamma As Variant, theta As Variant, vega As Variant
            Dim strat As String

            Select Case kind
                Case "fut"
                    typStr = "Future"
                    undl = wsR.Cells(i, 7).Value                   ' G PRODUCT
                    undlPx = wsR.Cells(i, 12).Value                ' L LAST
                    strike = ""
                    chgPct = wsR.Cells(i, 13).Value                ' M % CHG
                    delta = wsR.Cells(i, 18).Value                 ' R DELTA
                    gamma = "": theta = "": vega = ""
                    strat = "Futures"
                Case "opt"
                    typStr = CStr(wsR.Cells(i, 8).Value)           ' H TYPE
                    undl = wsR.Cells(i, 10).Value                  ' J UNDERLYING
                    undlPx = wsR.Cells(i, 11).Value                ' K UNDL LAST
                    strike = wsR.Cells(i, 9).Value                 ' I STRIKE
                    chgPct = wsR.Cells(i, 16).Value                ' P % CHG
                    delta = wsR.Cells(i, 20).Value                 ' T DELTA
                    gamma = wsR.Cells(i, 21).Value                 ' U GAMMA
                    theta = wsR.Cells(i, 22).Value                 ' V THETA
                    vega = wsR.Cells(i, 23).Value                  ' W VEGA
                    strat = CStr(wsR.Cells(i, 28).Value)           ' AB STRATEGY TAG
                Case "wrt"
                    typStr = CStr(wsR.Cells(i, 7).Value)           ' G TYPE
                    undl = wsR.Cells(i, 9).Value                   ' I UNDERLYING
                    undlPx = wsR.Cells(i, 10).Value                ' J UNDL LAST
                    strike = wsR.Cells(i, 11).Value                ' K STRIKE
                    chgPct = wsR.Cells(i, 16).Value                ' P % CHG
                    delta = wsR.Cells(i, 20).Value                 ' T DELTA
                    gamma = wsR.Cells(i, 24).Value                 ' X GAMMA
                    theta = wsR.Cells(i, 25).Value                 ' Y THETA
                    vega = wsR.Cells(i, 26).Value                  ' Z VEGA
                    strat = "Warrant"
            End Select

            wsA.Cells(r, 3).Value = typStr
            wsA.Cells(r, 5).Value = undl
            wsA.Cells(r, 6).Value = undlPx
            wsA.Cells(r, 7).Value = strike
            wsA.Cells(r, 9).Value = chgPct
            wsA.Cells(r, 11).Value = delta
            wsA.Cells(r, 12).Value = gamma
            wsA.Cells(r, 13).Value = theta
            wsA.Cells(r, 14).Value = vega
            wsA.Cells(r, 15).Value = DeriveTrend(typStr, undlPx, strike, delta)
            wsA.Cells(r, 16).Value = strat
            wsA.Cells(r, 17).Value = DeriveNote(wsR.Cells(i, 5).Value, delta, chgPct)

            ' Number formats
            wsA.Cells(r, 6).NumberFormat = "#,##0.00"
            wsA.Cells(r, 7).NumberFormat = "#,##0.00"
            wsA.Cells(r, 9).NumberFormat = "0.00%"
            wsA.Cells(r, 11).NumberFormat = "0.0000"
            wsA.Cells(r, 12).NumberFormat = "0.000000"
            wsA.Cells(r, 13).NumberFormat = "0.0000"
            wsA.Cells(r, 14).NumberFormat = "0.0000"

            ' --- Bucket: DTE ---
            Dim dteBucket As Long: dteBucket = -1
            If IsNumeric(wsR.Cells(i, 5).Value) Then
                Dim dteVal As Long: dteVal = CLng(wsR.Cells(i, 5).Value)
                Select Case True
                    Case dteVal <= 7:  dteBucket = 0
                    Case dteVal <= 30: dteBucket = 1
                    Case dteVal <= 60: dteBucket = 2
                    Case Else:         dteBucket = 3
                End Select
                dteCnt(dteBucket) = dteCnt(dteBucket) + 1
                If IsNumeric(delta) Then dteAbsDelta(dteBucket) = dteAbsDelta(dteBucket) + Abs(CDbl(delta))
                If IsNumeric(theta) Then dteTheta(dteBucket) = dteTheta(dteBucket) + CDbl(theta)
            End If

            ' --- Bucket: Strategy ---
            If strat <> "" Then
                If Not stratDict.Exists(strat) Then stratDict(strat) = Array(0&, 0#, 0#)
                Dim sArr As Variant: sArr = stratDict(strat)
                sArr(0) = sArr(0) + 1
                If IsNumeric(delta) Then sArr(1) = sArr(1) + CDbl(delta)
                If IsNumeric(theta) Then sArr(2) = sArr(2) + CDbl(theta)
                stratDict(strat) = sArr
            End If

            ' --- Bucket: Underlying ---
            Dim uKey As String: uKey = Trim(CStr(undl))
            If uKey <> "" Then
                If Not undlDict.Exists(uKey) Then undlDict(uKey) = Array(0&, 0#, 0#)
                Dim uArr As Variant: uArr = undlDict(uKey)
                uArr(0) = uArr(0) + 1
                If IsNumeric(delta) Then uArr(1) = uArr(1) + CDbl(delta)
                If IsNumeric(undlPx) Then uArr(2) = uArr(2) + CDbl(undlPx)
                undlDict(uKey) = uArr
            End If

            r = r + 1
NextRow:
        Next i
    Next sec

    ' =================================================================
    ' Top-area summary sections (col D / I / N) alongside KPI bar (col A-B)
    ' =================================================================

    ' --- DTE Bucket (cols D-G, rows 2-7) ---
    wsA.Cells(2, 4).Value = "DTE BUCKET"
    wsA.Cells(3, 4).Value = "Bucket"
    wsA.Cells(3, 5).Value = "Count"
    wsA.Cells(3, 6).Value = "Sum |Delta|"
    wsA.Cells(3, 7).Value = "Sum Theta/day"
    Dim labels As Variant: labels = Array("<=7d", "8-30d", "31-60d", ">60d")
    Dim b As Long
    For b = 0 To 3
        wsA.Cells(4 + b, 4).Value = labels(b)
        wsA.Cells(4 + b, 5).Value = dteCnt(b)
        wsA.Cells(4 + b, 6).Value = dteAbsDelta(b): wsA.Cells(4 + b, 6).NumberFormat = "0.0000"
        wsA.Cells(4 + b, 7).Value = dteTheta(b):    wsA.Cells(4 + b, 7).NumberFormat = "0.0000"
    Next b

    ' --- Strategy (cols I-L, rows 2 onwards) ---
    wsA.Cells(2, 9).Value = "STRATEGY"
    wsA.Cells(3, 9).Value = "Strategy"
    wsA.Cells(3, 10).Value = "Count"
    wsA.Cells(3, 11).Value = "Sum Delta"
    wsA.Cells(3, 12).Value = "Sum Theta/day"
    Dim sR As Long: sR = 4
    Dim sKey As Variant
    For Each sKey In stratDict.Keys
        Dim sV As Variant: sV = stratDict(sKey)
        wsA.Cells(sR, 9).Value = sKey
        wsA.Cells(sR, 10).Value = sV(0)
        wsA.Cells(sR, 11).Value = sV(1): wsA.Cells(sR, 11).NumberFormat = "0.0000"
        wsA.Cells(sR, 12).Value = sV(2): wsA.Cells(sR, 12).NumberFormat = "0.0000"
        sR = sR + 1
        If sR > 13 Then Exit For   ' max 10 strategies (rows 4-13)
    Next sKey

    ' --- Underlying (cols N-Q, rows 2 onwards) ---
    wsA.Cells(2, 14).Value = "UNDERLYING"
    wsA.Cells(3, 14).Value = "Underlying"
    wsA.Cells(3, 15).Value = "Pos Count"
    wsA.Cells(3, 16).Value = "Sum Delta"
    wsA.Cells(3, 17).Value = "Avg Spot"
    Dim uR As Long: uR = 4
    Dim uKeyV As Variant
    For Each uKeyV In undlDict.Keys
        Dim uV As Variant: uV = undlDict(uKeyV)
        wsA.Cells(uR, 14).Value = uKeyV
        wsA.Cells(uR, 15).Value = uV(0)
        wsA.Cells(uR, 16).Value = uV(1): wsA.Cells(uR, 16).NumberFormat = "0.0000"
        If uV(0) > 0 Then
            wsA.Cells(uR, 17).Value = uV(2) / uV(0)
            wsA.Cells(uR, 17).NumberFormat = "#,##0.00"
        End If
        uR = uR + 1
        If uR > 10 Then Exit For   ' max 7 underlyings (rows 4-10)
    Next uKeyV

    Exit Sub
Fail:
    MsgBox "RefreshAnalysis failed: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' SCENARIO ANALYSIS (F1)
' ---------------------------------------------------------------------
' Re-evaluates every option / warrant leg under bumped spot / IV scenarios
' using Black-Scholes. Aggregates portfolio-level P&L impact and writes
' a small grid to Analysis worksheet rows 40-46.
'
' Scenarios:
'   Spot bumps: -5% / -3% / -1% / 0 / +1% / +3% / +5%
'   IV bumps:   -5pp / 0 / +5pp  (absolute, not relative; e.g. 0.20 -> 0.25)
'
' Output layout (Analysis sheet):
'   Row 40  "SCENARIO ANALYSIS (Portfolio P&L, TWD)"
'   Row 41  Header:  "Spot Bump" | "-5%" | "-3%" | "-1%" | "0" | "+1%" | "+3%" | "+5%" | (sep) | "IV-5pp" | "IV+5pp"
'   Row 42  PnL row:  <name> | val | val | val | val | val | val | val | (sep) | val | val
'
' Futures contribute via direct spot delta (linear, 1:1 P&L per point per contract * mult).
' Equity/unknown types skipped.
'
' Errors at each stage are caught and reported with stage tag for
' rapid diagnosis (e.g. "F1 Stage 2 read positions: ...").
' =====================================================================
Public Sub BuildScenarioAnalysis()
    Dim stage As String: stage = "init"
    On Error GoTo Fail

    stage = "open sheets"
    Dim wsA As Worksheet, wsR As Worksheet, wsC As Worksheet
    Set wsA = ThisWorkbook.Sheets("Analysis")
    Set wsR = ThisWorkbook.Sheets("RR5")
    Set wsC = ThisWorkbook.Sheets("_Config")

    stage = "read risk-free rate"
    Dim rRate As Double
    rRate = ToDblD(GetConfigVal(wsC, "BS_RISK_FREE_RATE"))
    If rRate <= 0 Then rRate = 0.02

    stage = "define scenarios"
    Dim spotBumps As Variant: spotBumps = Array(-0.05, -0.03, -0.01, 0#, 0.01, 0.03, 0.05)
    Dim ivBumps As Variant: ivBumps = Array(-0.05, 0.05)   ' absolute pp

    stage = "clear old output rows 40-50"
    wsA.Range("A40:Q50").ClearContents

    stage = "write headers"
    wsA.Cells(40, 1).Value = "SCENARIO ANALYSIS (Portfolio P&L, TWD)"
    wsA.Cells(41, 1).Value = "Spot Bump"
    Dim h As Long
    For h = 0 To UBound(spotBumps)
        wsA.Cells(41, 2 + h).Value = ScenarioLabel(CDbl(spotBumps(h)), "pct")
    Next h
    wsA.Cells(41, 2 + UBound(spotBumps) + 1).Value = "|"
    For h = 0 To UBound(ivBumps)
        wsA.Cells(41, 2 + UBound(spotBumps) + 2 + h).Value = ScenarioLabel(CDbl(ivBumps(h)), "ivpp")
    Next h

    ' Initialize accumulator arrays
    Dim nSpot As Long: nSpot = UBound(spotBumps) + 1
    Dim nIv As Long: nIv = UBound(ivBumps) + 1
    Dim pnlSpot() As Double: ReDim pnlSpot(0 To nSpot - 1)
    Dim pnlIV() As Double: ReDim pnlIV(0 To nIv - 1)

    stage = "scan Options rows 19-26"
    Call ScanScenarioOptionsSection(wsR, 19, 26, rRate, spotBumps, ivBumps, pnlSpot, pnlIV)

    stage = "scan Warrants rows 31-38"
    Call ScanScenarioWarrantsSection(wsR, 31, 38, rRate, spotBumps, ivBumps, pnlSpot, pnlIV)

    stage = "scan Futures rows 7-14 (linear delta only)"
    Call ScanScenarioFuturesSection(wsR, 7, 14, spotBumps, pnlSpot)

    stage = "write portfolio P&L row"
    wsA.Cells(42, 1).Value = "Portfolio P&L"
    Dim sc As Long
    For sc = 0 To nSpot - 1
        wsA.Cells(42, 2 + sc).Value = pnlSpot(sc)
        wsA.Cells(42, 2 + sc).NumberFormat = "#,##0;[Red](#,##0)"
    Next sc
    For sc = 0 To nIv - 1
        wsA.Cells(42, 2 + nSpot + 1 + sc).Value = pnlIV(sc)
        wsA.Cells(42, 2 + nSpot + 1 + sc).NumberFormat = "#,##0;[Red](#,##0)"
    Next sc

    stage = "done"
    Application.StatusBar = "Scenario Analysis OK " & Format(Now, "hh:mm:ss")
    Exit Sub
Fail:
    Application.StatusBar = False
    MsgBox "BuildScenarioAnalysis failed at stage [" & stage & "]:" & vbCrLf & _
           Err.Description, vbExclamation, "F1 Scenario Analysis"
End Sub

' --- Helpers ----------------------------------------------------------

Private Sub ScanScenarioOptionsSection(wsR As Worksheet, rStart As Long, rEnd As Long, _
    rRate As Double, spotBumps As Variant, ivBumps As Variant, _
    ByRef pnlSpot() As Double, ByRef pnlIV() As Double)

    Dim stage As String: stage = "opt init"
    On Error GoTo Fail
    Dim i As Long
    For i = rStart To rEnd
        stage = "opt row " & i & " read"
        Dim contract As String: contract = Trim(CStr(wsR.Cells(i, 1).Value))
        If contract = "" Then GoTo NextI

        Dim S As Double:   S = ToDblD(wsR.Cells(i, 11).Value)   ' K UNDL LAST
        Dim K As Double:   K = ToDblD(wsR.Cells(i, 9).Value)    ' I STRIKE
        Dim dte As Double: dte = ToDblD(wsR.Cells(i, 5).Value)  ' E DTE
        Dim iv As Double:  iv = ToDblD(wsR.Cells(i, 19).Value)  ' S IV
        Dim signedContracts As Double: signedContracts = ToDblD(wsR.Cells(i, 12).Value)  ' L
        Dim mult As Double: mult = ToDblD(wsR.Cells(i, 13).Value)                        ' M
        Dim oType As String: oType = CStr(wsR.Cells(i, 8).Value)                         ' H
        Dim curPrem As Double: curPrem = ToDblD(wsR.Cells(i, 15).Value)                  ' O LAST PREM

        If S <= 0 Or K <= 0 Or dte <= 0 Or iv <= 0 Then GoTo NextI
        If signedContracts = 0 Or mult = 0 Then GoTo NextI

        Dim T As Double: T = dte / 365#

        stage = "opt row " & i & " baseline price"
        Dim basePx As Double
        basePx = BS_Price(S, K, T, rRate, iv, oType)
        ' If LAST PREM is filled, use it as baseline (more accurate). Else BS theoretical.
        Dim baseline As Double
        If curPrem > 0 Then baseline = curPrem Else baseline = basePx

        ' --- spot scenarios ---
        stage = "opt row " & i & " spot scenarios"
        Dim sb As Long
        For sb = 0 To UBound(spotBumps)
            Dim Sn As Double: Sn = S * (1 + CDbl(spotBumps(sb)))
            Dim newPx As Double: newPx = BS_Price(Sn, K, T, rRate, iv, oType)
            Dim deltaPx As Double: deltaPx = newPx - basePx   ' use BS-based delta to avoid baseline noise
            pnlSpot(sb) = pnlSpot(sb) + signedContracts * mult * deltaPx
        Next sb

        ' --- iv scenarios ---
        stage = "opt row " & i & " iv scenarios"
        Dim ib As Long
        For ib = 0 To UBound(ivBumps)
            Dim ivN As Double: ivN = iv + CDbl(ivBumps(ib))
            If ivN < 0.001 Then ivN = 0.001
            Dim ivPx As Double: ivPx = BS_Price(S, K, T, rRate, ivN, oType)
            pnlIV(ib) = pnlIV(ib) + signedContracts * mult * (ivPx - basePx)
        Next ib
NextI:
    Next i
    Exit Sub
Fail:
    Err.Raise Err.Number, "ScanScenarioOptionsSection / " & stage, Err.Description
End Sub

Private Sub ScanScenarioWarrantsSection(wsR As Worksheet, rStart As Long, rEnd As Long, _
    rRate As Double, spotBumps As Variant, ivBumps As Variant, _
    ByRef pnlSpot() As Double, ByRef pnlIV() As Double)

    Dim stage As String: stage = "wrt init"
    On Error GoTo Fail
    Dim i As Long
    For i = rStart To rEnd
        stage = "wrt row " & i & " read"
        Dim code As String: code = Trim(CStr(wsR.Cells(i, 1).Value))
        If code = "" Then GoTo NextI

        Dim S As Double: S = ToDblD(wsR.Cells(i, 10).Value)     ' J UNDL LAST
        Dim K As Double: K = ToDblD(wsR.Cells(i, 11).Value)     ' K STRIKE
        Dim conv As Double: conv = ToDblD(wsR.Cells(i, 12).Value) ' L CONV RATIO
        Dim shares As Double: shares = ToDblD(wsR.Cells(i, 13).Value)  ' M
        Dim dte As Double: dte = ToDblD(wsR.Cells(i, 5).Value)
        Dim iv As Double: iv = ToDblD(wsR.Cells(i, 19).Value)
        Dim wType As String: wType = CStr(wsR.Cells(i, 7).Value)

        If S <= 0 Or K <= 0 Or conv <= 0 Or shares = 0 Or dte <= 0 Or iv <= 0 Then GoTo NextI

        Dim T As Double: T = dte / 365#
        Dim oTypeW As String
        If InStr(UCase(wType), "CALL") > 0 Or InStr(UCase(wType), "BULL") > 0 Then
            oTypeW = "Call"
        Else
            oTypeW = "Put"
        End If

        stage = "wrt row " & i & " baseline"
        Dim baseW As Double: baseW = BS_Price(S, K, T, rRate, iv, oTypeW) / conv

        stage = "wrt row " & i & " spot bumps"
        Dim sb As Long
        For sb = 0 To UBound(spotBumps)
            Dim Sn As Double: Sn = S * (1 + CDbl(spotBumps(sb)))
            Dim newW As Double: newW = BS_Price(Sn, K, T, rRate, iv, oTypeW) / conv
            pnlSpot(sb) = pnlSpot(sb) + shares * (newW - baseW)
        Next sb

        stage = "wrt row " & i & " iv bumps"
        Dim ib As Long
        For ib = 0 To UBound(ivBumps)
            Dim ivN As Double: ivN = iv + CDbl(ivBumps(ib))
            If ivN < 0.001 Then ivN = 0.001
            Dim ivW As Double: ivW = BS_Price(S, K, T, rRate, ivN, oTypeW) / conv
            pnlIV(ib) = pnlIV(ib) + shares * (ivW - baseW)
        Next ib
NextI:
    Next i
    Exit Sub
Fail:
    Err.Raise Err.Number, "ScanScenarioWarrantsSection / " & stage, Err.Description
End Sub

Private Sub ScanScenarioFuturesSection(wsR As Worksheet, rStart As Long, rEnd As Long, _
    spotBumps As Variant, ByRef pnlSpot() As Double)

    Dim stage As String: stage = "fut init"
    On Error GoTo Fail
    Dim i As Long
    For i = rStart To rEnd
        stage = "fut row " & i
        Dim contract As String: contract = Trim(CStr(wsR.Cells(i, 1).Value))
        If contract = "" Then GoTo NextI

        Dim S As Double: S = ToDblD(wsR.Cells(i, 12).Value)    ' L LAST
        Dim contracts As Double: contracts = ToDblD(wsR.Cells(i, 9).Value)   ' I CONTRACTS
        Dim mult As Double: mult = ToDblD(wsR.Cells(i, 10).Value)            ' J MULTIPLIER
        Dim direction As String: direction = CStr(wsR.Cells(i, 8).Value)     ' H DIRECTION

        If S <= 0 Or contracts <= 0 Or mult <= 0 Then GoTo NextI
        Dim sign As Long: sign = IIf(UCase(direction) = "LONG", 1, -1)

        Dim sb As Long
        For sb = 0 To UBound(spotBumps)
            Dim deltaS As Double: deltaS = S * CDbl(spotBumps(sb))
            pnlSpot(sb) = pnlSpot(sb) + sign * contracts * mult * deltaS
        Next sb
NextI:
    Next i
    Exit Sub
Fail:
    Err.Raise Err.Number, "ScanScenarioFuturesSection / " & stage, Err.Description
End Sub

Private Function ScenarioLabel(v As Double, kind As String) As String
    If kind = "pct" Then
        If v = 0 Then ScenarioLabel = "0" Else ScenarioLabel = Format(v, "+0%;-0%")
    Else
        If v = 0 Then ScenarioLabel = "IV 0" Else ScenarioLabel = "IV" & Format(v * 100, "+0;-0") & "pp"
    End If
End Function

Private Function GetConfigVal(wsC As Worksheet, key As String) As Variant
    Dim i As Long
    For i = 1 To 30
        If wsC.Cells(i, 1).Value = key Then
            GetConfigVal = wsC.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    GetConfigVal = ""
End Function

Private Function ToDblD(v As Variant) As Double
    If IsNumeric(v) Then ToDblD = CDbl(v) Else ToDblD = 0
End Function

' =====================================================================
' END SCENARIO ANALYSIS (F1)
' =====================================================================

' =====================================================================
' F4: EQUITY CURVE + DRAWDOWN CHART
' ---------------------------------------------------------------------
' Reads HistoryLog (cols A=Date, B=TotalMarketValue), writes two
' chart-helper data tables to Analysis worksheet rows 55-80, then
' creates/updates two ChartObjects: "EquityCurve" + "DrawdownChart".
'
' Drawdown formula:  dd_t = (MV_t - peak_t) / peak_t  where peak_t =
' running max of MV up to time t. Always <= 0.
'
' Output layout (Analysis):
'   Row 55  "EQUITY CURVE DATA"
'   Row 56  Date | Total MV | Drawdown%
'   Row 57+ time series
'
' Charts anchored to D52 (Equity) and L52 (Drawdown).
' =====================================================================
Public Sub BuildEquityDrawdownCharts()
    Dim stage As String: stage = "init"
    On Error GoTo Fail

    stage = "open sheets"
    Dim wsA As Worksheet, wsH As Worksheet
    Set wsA = ThisWorkbook.Sheets("Analysis")
    Set wsH = ThisWorkbook.Sheets("HistoryLog")

    stage = "read HistoryLog"
    Dim lr As Long: lr = wsH.Cells(wsH.Rows.Count, "A").End(xlUp).Row
    If lr < 3 Then
        MsgBox "HistoryLog has fewer than 2 data rows -- cannot draw charts.", _
               vbInformation, "F4 Equity Curve"
        Exit Sub
    End If

    stage = "clear old data table"
    wsA.Range("A55:C200").ClearContents
    wsA.Cells(55, 1).Value = "EQUITY CURVE DATA"
    wsA.Cells(56, 1).Value = "Date"
    wsA.Cells(56, 2).Value = "Total MV"
    wsA.Cells(56, 3).Value = "Drawdown%"

    stage = "iterate HistoryLog and compute drawdown"
    Dim peak As Double: peak = 0
    Dim outR As Long: outR = 57
    Dim i As Long
    For i = 2 To lr
        Dim dt As Variant: dt = wsH.Cells(i, 1).Value
        Dim mv As Variant: mv = wsH.Cells(i, 2).Value
        If Not IsDate(dt) Then GoTo NextI
        If Not IsNumeric(mv) Then GoTo NextI
        If CDbl(mv) <= 0 Then GoTo NextI

        If CDbl(mv) > peak Then peak = CDbl(mv)
        Dim dd As Double
        If peak > 0 Then dd = (CDbl(mv) - peak) / peak Else dd = 0

        wsA.Cells(outR, 1).Value = CDate(dt)
        wsA.Cells(outR, 1).NumberFormat = "yyyy/m/d"
        wsA.Cells(outR, 2).Value = CDbl(mv)
        wsA.Cells(outR, 2).NumberFormat = "#,##0"
        wsA.Cells(outR, 3).Value = dd
        wsA.Cells(outR, 3).NumberFormat = "0.00%"
        outR = outR + 1
NextI:
    Next i
    Dim lastR As Long: lastR = outR - 1
    If lastR < 58 Then
        MsgBox "Not enough valid HistoryLog rows.", vbInformation
        Exit Sub
    End If

    stage = "build/update Equity Curve chart"
    Call EnsureChart(wsA, "EquityCurve", "D52", _
        wsA.Range(wsA.Cells(57, 1), wsA.Cells(lastR, 1)), _
        wsA.Range(wsA.Cells(57, 2), wsA.Cells(lastR, 2)), _
        "Equity Curve  (Total MV)", "Date", "Total MV (TWD)", _
        Nothing, Nothing, "")

    stage = "build/update Drawdown chart"
    Call EnsureChart(wsA, "DrawdownChart", "L52", _
        wsA.Range(wsA.Cells(57, 1), wsA.Cells(lastR, 1)), _
        wsA.Range(wsA.Cells(57, 3), wsA.Cells(lastR, 3)), _
        "Drawdown Underwater", "Date", "Drawdown %", _
        Nothing, Nothing, "")

    Application.StatusBar = "Equity + Drawdown charts updated"
    Exit Sub
Fail:
    Application.StatusBar = False
    MsgBox "BuildEquityDrawdownCharts failed at stage [" & stage & "]:" & vbCrLf & _
           Err.Description, vbExclamation, "F4 Equity Curve"
End Sub

' --- Helper: find or create a chart anchored to a cell, with 1 series ---
Private Sub EnsureChart(wsA As Worksheet, chartName As String, anchorCell As String, _
                         xRange As Range, yRange As Range, _
                         title As String, xLabel As String, yLabel As String, _
                         Optional zX As Variant, Optional zY As Variant, _
                         Optional extraTitle As String = "")
    Dim co As ChartObject, cc As ChartObject
    Dim found As Boolean: found = False
    For Each cc In wsA.ChartObjects
        If cc.Name = chartName Then
            Set co = cc
            found = True
            Exit For
        End If
    Next cc
    If Not found Then
        Set co = wsA.ChartObjects.Add( _
            Left:=wsA.Range(anchorCell).Left, _
            Top:=wsA.Range(anchorCell).Top, _
            Width:=460, Height:=260)
        co.Name = chartName
    End If

    With co.Chart
        .ChartType = xlXYScatterLines
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Dim sr As Series: Set sr = .SeriesCollection.NewSeries
        sr.XValues = xRange
        sr.Values = yRange
        sr.Name = chartName

        .HasTitle = True
        .ChartTitle.Text = title
        .HasLegend = False
        On Error Resume Next
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = xLabel
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = yLabel
        On Error GoTo 0
    End With
End Sub
' =====================================================================
' END F4
' =====================================================================

Private Function DeriveTrend(typStr As String, undlPx As Variant, _
                              strike As Variant, delta As Variant) As String
    If Not IsNumeric(undlPx) Or Not IsNumeric(strike) Then
        DeriveTrend = "-"
        Exit Function
    End If
    Dim t As String: t = LCase(typStr)
    Dim isCall As Boolean
    isCall = (InStr(t, "call") > 0 Or InStr(t, "bull") > 0)
    Dim u As Double: u = CDbl(undlPx)
    Dim k As Double: k = CDbl(strike)
    Dim m As String

    If isCall Then
        If u > k * 1.02 Then
            m = "ITM"
        ElseIf u < k * 0.98 Then
            m = "OTM"
        Else
            m = "ATM"
        End If
    Else
        If u < k * 0.98 Then
            m = "ITM"
        ElseIf u > k * 1.02 Then
            m = "OTM"
        Else
            m = "ATM"
        End If
    End If

    DeriveTrend = m & " " & IIf(isCall, "Bull", "Bear")
End Function

Private Function DeriveNote(dte As Variant, delta As Variant, chgPct As Variant) As String
    Dim notes As String: notes = ""
    If IsNumeric(dte) Then
        If CLng(dte) <= 7 And CLng(dte) > 0 Then
            notes = notes & "EXP<=7d; "
        ElseIf CLng(dte) <= 30 Then
            notes = notes & "<=30d; "
        End If
    End If
    If IsNumeric(delta) Then
        Dim d As Double: d = Abs(CDbl(delta))
        If d > 0.8 Then
            notes = notes & "Deep ITM; "
        ElseIf d < 0.15 And d > 0 Then
            notes = notes & "Far OTM; "
        End If
    End If
    If IsNumeric(chgPct) Then
        Dim c As Double: c = CDbl(chgPct)
        If c > 0.5 Then
            notes = notes & "+" & Format(c, "0%") & "; "
        ElseIf c < -0.3 Then
            notes = notes & Format(c, "0%") & "; "
        End If
    End If
    If Right(notes, 2) = "; " Then notes = Left(notes, Len(notes) - 2)
    DeriveNote = notes
End Function

' =====================================================================
' Label writing (idempotent)
' =====================================================================
Private Sub EnsureDashboardLabels(wsR As Worksheet)
    Dim row3Lbl As Variant
    row3Lbl = Array( _
        Array("A3", "IV RANK"), Array("C3", "IV PCTL"), _
        Array("E3", "THETA CUM"), Array("G3", "DELTA D7"), _
        Array("I3", "EXP <=7D"), Array("K3", "AVG DTE"), _
        Array("M3", "TRADES"), Array("O3", "30D TRD"), _
        Array("Q3", "WIN%"), Array("S3", "AVG WIN"), _
        Array("U3", "AVG LOSS"), Array("W3", "PF"))

    Dim row4Lbl As Variant
    row4Lbl = Array( _
        Array("A4", "MV 7D"), Array("C4", "MV 30D"), _
        Array("E4", "MV CHG7%"), Array("G4", "MV CHG30%"), _
        Array("I4", "MDD%"), Array("K4", "MEAN DD%"), _
        Array("M4", "BEST D%"), Array("O4", "WORST D%"), _
        Array("Q4", "DAYS LOG"), Array("S4", "IV 30D"))

    Dim i As Long
    For i = LBound(row3Lbl) To UBound(row3Lbl)
        With wsR.Range(row3Lbl(i)(0))
            .Value = row3Lbl(i)(1)
            .Font.Bold = True
            .Font.Color = RGB(150, 150, 150)
            .Font.Size = 9
            .Font.Name = "Consolas"
        End With
    Next i
    For i = LBound(row4Lbl) To UBound(row4Lbl)
        With wsR.Range(row4Lbl(i)(0))
            .Value = row4Lbl(i)(1)
            .Font.Bold = True
            .Font.Color = RGB(150, 150, 150)
            .Font.Size = 9
            .Font.Name = "Consolas"
        End With
    Next i
End Sub

' =====================================================================
' Realized aggregates: Trades / Win% / Avg Win / Avg Loss / Profit Factor
'   Realized cols: A=TICKER B=TYPE C=STRATEGY D=PNL E=SHARES F=AVG COST
'                  G=NET AMT H=STRATEGY TAG I=PNL(TWD) J=DATE
' =====================================================================
Private Sub WriteRealizedMetrics(wsR As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Realized")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then GoTo WriteOut

    Dim totalTrades As Long, wins As Long, losses As Long
    Dim sumWin As Double, sumLoss As Double, totalPnL As Double
    Dim trades30D As Long
    Dim cutDate As Date: cutDate = Date - 30

    Dim i As Long
    For i = 2 To lr
        Dim pnl As Variant: pnl = ws.Cells(i, 9).Value   ' I = PNL(TWD)
        Dim dt As Variant:  dt = ws.Cells(i, 10).Value   ' J = DATE
        If IsNumeric(pnl) Then
            totalTrades = totalTrades + 1
            totalPnL = totalPnL + CDbl(pnl)
            If CDbl(pnl) > 0 Then
                wins = wins + 1
                sumWin = sumWin + CDbl(pnl)
            ElseIf CDbl(pnl) < 0 Then
                losses = losses + 1
                sumLoss = sumLoss + CDbl(pnl)
            End If
            If IsDate(dt) Then
                If CDate(dt) >= cutDate Then trades30D = trades30D + 1
            End If
        End If
    Next i

WriteOut:
    With wsR
        .Range("N3").Value = totalTrades
        .Range("P3").Value = trades30D
        If totalTrades > 0 Then
            .Range("R3").Value = wins / totalTrades
            .Range("R3").NumberFormat = "0%"
        Else
            .Range("R3").Value = 0
        End If
        If wins > 0 Then
            .Range("T3").Value = sumWin / wins
            .Range("T3").NumberFormat = "#,##0"
        End If
        If losses > 0 Then
            .Range("V3").Value = sumLoss / losses
            .Range("V3").NumberFormat = "#,##0;(#,##0)"
        End If
        If sumLoss <> 0 Then
            .Range("X3").Value = Abs(sumWin / sumLoss)
            .Range("X3").NumberFormat = "0.00"
        End If
    End With
End Sub

' =====================================================================
' IVTracker aggregates: latest IV Rank / Percentile per underlying, avg
'   IVTracker cols: A=Date B=Underlying C=IV Current D=IV HV30
'                   E=IV Percentile F=IV Rank
' =====================================================================
Private Sub WriteIVMetrics(wsR As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("IVTracker")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Exit Sub

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim entry(3) As Variant
    Dim i As Long
    For i = 2 To lr
        Dim dt As Variant: dt = ws.Cells(i, 1).Value
        Dim und As String: und = Trim(CStr(ws.Cells(i, 2).Value))
        If und = "" Or Not IsDate(dt) Then GoTo NextI

        Dim shouldUpdate As Boolean: shouldUpdate = True
        If dict.Exists(und) Then
            Dim existing As Variant: existing = dict(und)
            If CDate(existing(0)) >= CDate(dt) Then shouldUpdate = False
        End If

        If shouldUpdate Then
            entry(0) = dt
            entry(1) = ws.Cells(i, 3).Value   ' IV Current
            entry(2) = ws.Cells(i, 5).Value   ' IV Percentile
            entry(3) = ws.Cells(i, 6).Value   ' IV Rank
            dict(und) = entry
        End If
NextI:
    Next i

    Dim sumPctl As Double, sumRank As Double, cnt As Long
    Dim key As Variant
    For Each key In dict.Keys
        Dim ent As Variant: ent = dict(key)
        If IsNumeric(ent(2)) Then sumPctl = sumPctl + CDbl(ent(2))
        If IsNumeric(ent(3)) Then sumRank = sumRank + CDbl(ent(3))
        cnt = cnt + 1
    Next key

    With wsR
        If cnt > 0 Then
            .Range("B3").Value = sumRank / cnt
            .Range("B3").NumberFormat = "0.0"
            .Range("D3").Value = sumPctl / cnt
            .Range("D3").NumberFormat = "0.0"
        End If
    End With
End Sub

' =====================================================================
' HistoryLog metrics: MV time-series, MDD, best/worst day
'   HistoryLog cols: A=Date B=TotalMV C=CumPnL D-G=NetGreeks
'                    H=Px(TX) I=Ret%(TX) J=Px(SPY) K=Ret%(SPY)
'                    L=RealizedPnL M=PortRet% N=PortRetChg% O=IV_Avg
' =====================================================================
Private Sub WriteHistoryLogMetrics(wsR As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HistoryLog")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Exit Sub

    Dim today As Date: today = Date
    Dim cut7 As Date:  cut7 = today - 7
    Dim cut30 As Date: cut30 = today - 30

    Dim curMV As Double, mv7d As Double, mv30d As Double
    Dim foundLatest As Boolean, found7 As Boolean, found30 As Boolean

    Dim i As Long
    For i = lr To 2 Step -1
        Dim dt As Variant: dt = ws.Cells(i, 1).Value
        Dim mv As Variant: mv = ws.Cells(i, 2).Value
        If Not IsDate(dt) Then GoTo NextRev
        If Not IsNumeric(mv) Then GoTo NextRev
        If Not foundLatest Then
            curMV = CDbl(mv): foundLatest = True
        End If
        If Not found7 And CDate(dt) <= cut7 Then
            mv7d = CDbl(mv): found7 = True
        End If
        If Not found30 And CDate(dt) <= cut30 Then
            mv30d = CDbl(mv): found30 = True
        End If
        If foundLatest And found7 And found30 Then Exit For
NextRev:
    Next i

    ' MDD / mean DD / best/worst day from PortRetChg% (col N)
    Dim peak As Double, dd As Double, maxDD As Double
    Dim sumDD As Double, ddCnt As Long
    Dim bestDay As Double, worstDay As Double
    Dim daysCnt As Long
    Dim ivSum As Double, ivCnt As Long
    bestDay = -1: worstDay = 1: peak = 0

    For i = 2 To lr
        Dim dti As Variant: dti = ws.Cells(i, 1).Value
        Dim mvi As Variant: mvi = ws.Cells(i, 2).Value
        Dim chg As Variant: chg = ws.Cells(i, 14).Value
        Dim ivi As Variant: ivi = ws.Cells(i, 15).Value

        If IsDate(dti) Then daysCnt = daysCnt + 1

        If IsNumeric(mvi) Then
            If CDbl(mvi) > peak Then peak = CDbl(mvi)
            If peak > 0 Then
                dd = (CDbl(mvi) - peak) / peak
                If dd < maxDD Then maxDD = dd
                If dd < 0 Then
                    sumDD = sumDD + dd
                    ddCnt = ddCnt + 1
                End If
            End If
        End If

        If IsNumeric(chg) Then
            If CDbl(chg) > bestDay Then bestDay = CDbl(chg)
            If CDbl(chg) < worstDay Then worstDay = CDbl(chg)
        End If

        If IsDate(dti) And IsNumeric(ivi) Then
            If CDate(dti) >= cut30 Then
                ivSum = ivSum + CDbl(ivi)
                ivCnt = ivCnt + 1
            End If
        End If
    Next i

    With wsR
        If mv7d <> 0 Then
            .Range("B4").Value = mv7d
            .Range("B4").NumberFormat = "#,##0"
            If mv7d > 0 And foundLatest Then
                .Range("F4").Value = (curMV - mv7d) / mv7d
                .Range("F4").NumberFormat = "0.00%"
            End If
        End If
        If mv30d <> 0 Then
            .Range("D4").Value = mv30d
            .Range("D4").NumberFormat = "#,##0"
            If mv30d > 0 And foundLatest Then
                .Range("H4").Value = (curMV - mv30d) / mv30d
                .Range("H4").NumberFormat = "0.00%"
            End If
        End If
        .Range("J4").Value = maxDD
        .Range("J4").NumberFormat = "0.00%"
        If ddCnt > 0 Then
            .Range("L4").Value = sumDD / ddCnt
            .Range("L4").NumberFormat = "0.00%"
        End If
        If bestDay > -1 Then
            .Range("N4").Value = bestDay
            .Range("N4").NumberFormat = "0.00%"
        End If
        If worstDay < 1 Then
            .Range("P4").Value = worstDay
            .Range("P4").NumberFormat = "0.00%"
        End If
        .Range("R4").Value = daysCnt
        If ivCnt > 0 Then
            .Range("T4").Value = ivSum / ivCnt
            .Range("T4").NumberFormat = "0.0%"
        End If
    End With
End Sub

' =====================================================================
' Greeks_History aggregates: cumulative Theta_PnL_Est, 7D Delta drift
'   Greeks_History cols: A=Date B=Spot_TX C=Port_Delta D=Port_Gamma
'                        E=Port_Theta F=Port_Vega G=IV_Avg
'                        H=Delta_PnL_Est I=Gamma_PnL_Est J=Theta_PnL_Est
'                        K=Vega_PnL_Est L=Residual_PnL
' =====================================================================
Private Sub WriteGreeksHistoryMetrics(wsR As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Greeks_History")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Exit Sub

    Dim thetaCum As Double
    Dim curDelta As Double, delta7d As Double
    Dim foundLatest As Boolean, found7 As Boolean
    Dim cut7 As Date: cut7 = Date - 7

    Dim i As Long
    For i = 2 To lr
        Dim t As Variant: t = ws.Cells(i, 10).Value   ' J = Theta_PnL_Est
        If IsNumeric(t) Then thetaCum = thetaCum + CDbl(t)
    Next i

    For i = lr To 2 Step -1
        Dim dt As Variant: dt = ws.Cells(i, 1).Value
        Dim d As Variant: d = ws.Cells(i, 3).Value
        If IsDate(dt) And IsNumeric(d) Then
            If Not foundLatest Then
                curDelta = CDbl(d): foundLatest = True
            End If
            If Not found7 And CDate(dt) <= cut7 Then
                delta7d = CDbl(d): found7 = True
            End If
            If foundLatest And found7 Then Exit For
        End If
    Next i

    With wsR
        .Range("F3").Value = thetaCum
        .Range("F3").NumberFormat = "#,##0;(#,##0)"
        If foundLatest And found7 Then
            .Range("H3").Value = curDelta - delta7d
            .Range("H3").NumberFormat = "+0.00;-0.00;0.00"
        End If
    End With
End Sub

' =====================================================================
' Position metrics from RR5 itself: expiring count, avg DTE
' =====================================================================
Private Sub WritePositionMetrics(wsR As Worksheet)
    Dim cntExp As Long, sumDTE As Double, cntDTE As Long

    ' Compact layout: Futures 7-14 / Options 19-26 / Warrants 31-38
    Dim secs As Variant
    secs = Array(Array(7, 14), Array(19, 26), Array(31, 38))

    Dim r As Variant
    For Each r In secs
        Dim i As Long
        For i = CLng(r(0)) To CLng(r(1))
            Dim dte As Variant: dte = wsR.Cells(i, 5).Value
            If IsNumeric(dte) Then
                If CLng(dte) > 0 Then
                    sumDTE = sumDTE + CLng(dte)
                    cntDTE = cntDTE + 1
                    If CLng(dte) <= 7 Then cntExp = cntExp + 1
                End If
            End If
        Next i
    Next r

    With wsR
        .Range("J3").Value = cntExp
        If cntDTE > 0 Then
            .Range("L3").Value = sumDTE / cntDTE
            .Range("L3").NumberFormat = "0"
        End If
    End With
End Sub
