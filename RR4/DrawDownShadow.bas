Attribute VB_Name = "DrawDownShadow"
Option Explicit

' ================================================================
' DRAWDOWN SHADOW v2.0
' Builds drawdown shadow + pain-adjusted chart from HistoryLog!M
'
' v2 changes vs v1:
'   1. Detection rule:
'        v1: >= 2 consecutive negative-return days
'        v2: any 5-day rolling window with >= 4 negative days marks
'            all 5 days in that window as "in drawdown"
'   2. cumRet now SUMS ONLY THE NEGATIVE returns within a segment.
'      Positive bounce days inside the segment are zeroed out.
'   3. BUG FIX: the pain-adjusted shadowVal formula was missing a
'      '+' between the two terms (the '_' line-continuation in v1
'      silently dropped it). v2 restores it.
' ================================================================

Private Const SH_HIST    As String = "HistoryLog"
Private Const SH_OUT     As String = "DrawdownChart"
Private Const DATA_ROW   As Long = 3
Private Const PAIN_CAP   As Double = 5
Private Const PAIN_OVERFLOW As Double = 9999

' Detection window parameters - change these to tune rule.
Private Const WIN_SIZE     As Long = 5    ' rolling window length
Private Const WIN_MIN_NEG  As Long = 4    ' min negative days within window

Public Sub BuildDrawdownShadow()
    Application.ScreenUpdating = False
    Application.StatusBar = "Building drawdown shadow v2..."
    On Error GoTo CleanFail

    ' --- Read HistoryLog!M into typed arrays ---
    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = ThisWorkbook.Sheets(SH_HIST): On Error GoTo CleanFail
    If wsH Is Nothing Then MsgBox "HistoryLog sheet not found", vbExclamation: GoTo CleanExit

    Dim lastRow As Long
    lastRow = wsH.Cells(wsH.Rows.Count, "M").End(xlUp).Row
    If lastRow < 3 Then MsgBox "Not enough data in column M (need >= 5 rows for 5-day window)", vbInformation: GoTo CleanExit

    Dim n As Long: n = lastRow - 1
    Dim ret() As Double: ReDim ret(1 To n)
    Dim dateStr() As String: ReDim dateStr(1 To n)
    Dim valid() As Boolean: ReDim valid(1 To n)

    Dim r As Long, i As Long
    For r = 2 To lastRow
        i = r - 1
        Dim v As Variant: v = wsH.Cells(r, "M").Value
        If IsNumeric(v) Then
            ret(i) = CDbl(v)
            valid(i) = True
        Else
            ret(i) = 0
            valid(i) = False
        End If
        dateStr(i) = "R" & r
        On Error Resume Next
        Dim d0 As Date: d0 = CDate(wsH.Cells(r, "A").Value)
        If d0 > 0 Then dateStr(i) = Format(d0, "m/d/yy")
        On Error GoTo CleanFail
    Next r

    ' --- Mark DD days using 5-day rolling window ---
    Dim isDD() As Boolean: ReDim isDD(1 To n)
    Dim w As Long, negCount As Long, k As Long
    For w = 1 To n - WIN_SIZE + 1
        negCount = 0
        For k = w To w + WIN_SIZE - 1
            If valid(k) And ret(k) < 0 Then negCount = negCount + 1
        Next k
        If negCount >= WIN_MIN_NEG Then
            For k = w To w + WIN_SIZE - 1
                isDD(k) = True
            Next k
        End If
    Next w

    ' --- Build segments from consecutive isDD-marked days ---
    Dim segments As Collection: Set segments = New Collection
    Dim segDates As Collection: Set segDates = New Collection

    Dim s As Long, segStart As Long, segEnd As Long
    s = 1
    Do While s <= n
        If isDD(s) Then
            segStart = s
            Do While s <= n
                If Not isDD(s) Then Exit Do
                s = s + 1
            Loop
            segEnd = s - 1

            Dim segLen As Long: segLen = segEnd - segStart + 1
            Dim segArr() As Double: ReDim segArr(0 To segLen - 1)
            Dim p As Long
            For p = 0 To segLen - 1
                ' Ignore positive returns - only negatives accumulate
                If ret(segStart + p) < 0 Then
                    segArr(p) = ret(segStart + p)
                Else
                    segArr(p) = 0
                End If
            Next p

            segments.Add segArr
            segDates.Add dateStr(segStart)
        Else
            s = s + 1
        End If
    Loop

    If segments.Count = 0 Then
        MsgBox "No drawdown segments found (no 5-day window has >= 4 negative days)", vbInformation
        GoTo CleanExit
    End If

    ' --- Find max segment length ---
    Dim maxDays As Long: maxDays = 0
    For s = 1 To segments.Count
        Dim a As Variant: a = segments(s)
        If UBound(a) + 1 > maxDays Then maxDays = UBound(a) + 1
    Next s

    ' --- Create / clear DrawdownChart sheet ---
    Dim wsD As Worksheet
    On Error Resume Next: Set wsD = ThisWorkbook.Sheets(SH_OUT): On Error GoTo CleanFail
    If wsD Is Nothing Then
        Set wsD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsD.Name = SH_OUT
    End If
    wsD.Cells.Clear

    Dim shp As Shape
    For Each shp In wsD.Shapes
        shp.Delete
    Next shp

    With wsD.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
    wsD.Activate
    ActiveWindow.DisplayGridlines = False

    With wsD.Cells(1, 1)
        .Value = "DRAWDOWN SHADOW v2 (rule: " & WIN_MIN_NEG & "/" & WIN_SIZE & ") | " & _
                 segments.Count & " episodes | " & Format(Now, "hh:mm:ss")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 12
    End With
    wsD.Rows(1).RowHeight = 24

    wsD.Cells(DATA_ROW - 1, 1).Value = "Day"
    wsD.Cells(DATA_ROW - 1, 1).Font.Color = RGB(150, 150, 150)

    Dim d As Long
    For d = 1 To maxDays
        wsD.Cells(DATA_ROW + d - 1, 1).Value = d
    Next d

    ' --- Compute shadow f(x) per segment per day ---
    Dim minCum As Double: minCum = 0
    Dim cumRet As Double, runMin As Double
    Dim xVal As Double, fxVal As Double, cellVal As Double
    Dim depth As Double, rVal As Long, colIdx As Long, segLenInner As Long

    For s = 1 To segments.Count
        Dim arr As Variant: arr = segments(s)
        segLenInner = UBound(arr) + 1
        colIdx = s + 1
        wsD.Cells(DATA_ROW - 1, colIdx).Value = segDates(s)
        wsD.Cells(DATA_ROW - 1, colIdx).Font.Color = RGB(180, 180, 180)
        wsD.Cells(DATA_ROW - 1, colIdx).Font.Size = 8

        cumRet = 0
        runMin = 0
        For d = 0 To segLenInner - 1
            xVal = d + 1
            cumRet = cumRet + arr(d)   ' arr(d) is 0 for positive days (ignored)
            fxVal = (Exp(xVal) / (xVal ^ 3 + WorksheetFunction.Pi() ^ 3) + Exp(Sqr(xVal))) * cumRet
            cellVal = fxVal
            wsD.Cells(DATA_ROW + d, colIdx).Value = Round(cellVal, 4)
            wsD.Cells(DATA_ROW + d, colIdx).NumberFormat = "0.0000"

            depth = Abs(cellVal)
            rVal = IIf(depth > 1000, 255, IIf(depth > 100, 230, IIf(depth > 10, 200, 170)))
            wsD.Cells(DATA_ROW + d, colIdx).Font.Color = RGB(rVal, 60, 60)

            If cellVal < runMin Then runMin = cellVal
        Next d
        If runMin < minCum Then minCum = runMin
    Next s

    ' --- Column widths / row heights ---
    wsD.Columns(1).ColumnWidth = 6
    Dim ci As Long
    For ci = 2 To segments.Count + 1
        wsD.Columns(ci).ColumnWidth = 9
    Next ci
    For d = 1 To maxDays
        wsD.Rows(DATA_ROW + d - 1).RowHeight = 15
    Next d

    ' --- Build chart ---
    Dim co As ChartObject
    Set co = wsD.ChartObjects.Add(wsD.Columns("M").Left, wsD.Rows(1).Top, 700, 400)
    With co.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Drawdown Shadow v2 (" & WIN_MIN_NEG & "/" & WIN_SIZE & " rule)"
        .ChartTitle.Font.Color = RGB(255, 192, 0)
        .ChartTitle.Font.Name = "Consolas"
        .ChartTitle.Font.Bold = True

        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        Dim xVals() As Long: ReDim xVals(0 To maxDays - 1)
        For d = 0 To maxDays - 1: xVals(d) = d + 1: Next d

        Dim yVals() As Variant
        Dim cum2 As Double, xv As Double
        Dim ser As Series
        Dim brightness As Long, lineClr As Long

        For s = 1 To segments.Count
            arr = segments(s)
            segLenInner = UBound(arr) + 1
            ReDim yVals(0 To maxDays - 1)
            cum2 = 0
            For d = 0 To maxDays - 1
                If d < segLenInner Then
                    xv = d + 1
                    cum2 = cum2 + arr(d)
                    yVals(d) = Round((Exp(xv) / (xv ^ 3 + WorksheetFunction.Pi() ^ 3) + Exp(Sqr(xv))) * cum2, 4)
                Else
                    yVals(d) = CVErr(xlErrNA)
                End If
            Next d
            Set ser = .SeriesCollection.NewSeries
            ser.Values = yVals
            ser.XValues = xVals
            ser.Name = CStr(segDates(s))

            brightness = 60 + CInt((s / segments.Count) * 160)
            lineClr = RGB(brightness, 20, 20)
            With ser.Format.Line
                .ForeColor.RGB = lineClr
                .Weight = IIf(s = segments.Count, 2, 1)
                .Transparency = IIf(s = segments.Count, 0, 0.4 - (s / segments.Count) * 0.3)
            End With
            ser.MarkerStyle = xlMarkerStyleNone
        Next s

        .PlotArea.Interior.Color = RGB(8, 8, 8)
        .PlotArea.Border.Color = RGB(40, 40, 40)
        .ChartArea.Interior.Color = RGB(0, 0, 0)
        .ChartArea.Border.Color = RGB(30, 30, 30)

        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Days in Drawdown"
            .AxisTitle.Font.Color = RGB(255, 255, 255)
            .AxisTitle.Font.Name = "Consolas"
            .TickLabels.Font.Color = RGB(255, 255, 255)
            .TickLabels.Font.Name = "Consolas"
            .TickLabels.Font.Size = 8
            .MajorGridlines.Border.Color = RGB(35, 35, 35)
        End With
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "f(x) = sum(neg) * (e^x/(x^3+pi^3) + e^sqrt(x))"
            .AxisTitle.Font.Color = RGB(255, 255, 255)
            .AxisTitle.Font.Name = "Consolas"
            .TickLabels.Font.Color = RGB(255, 255, 255)
            .TickLabels.Font.Name = "Consolas"
            .TickLabels.Font.Size = 8
            .TickLabels.NumberFormat = "0.00"
            .MajorGridlines.Border.Color = RGB(35, 35, 35)
            .MaximumScale = 0
            .CrossesAt = 0
        End With

        .HasLegend = True
        With .Legend
            .Interior.Color = RGB(10, 10, 10)
            .Border.Color = RGB(30, 30, 30)
            .Font.Color = RGB(255, 255, 255)
            .Font.Name = "Consolas"
            .Font.Size = 7
        End With
    End With

    ' ============================================================
    ' Pain-Adjusted Section  (BUG FIX: missing '+' restored)
    ' ============================================================
    Dim painStartRow As Long: painStartRow = DATA_ROW + maxDays + 2
    Dim painHeaderRow As Long: painHeaderRow = painStartRow + 1
    Dim painDataRow As Long
    Dim painVal As Double, shadowVal As Double
    Dim minPain As Double: minPain = 0

    With wsD.Cells(painStartRow, 1)
        .Value = "PAIN-ADJUSTED DRAWDOWN | f(x) = TAN(ASIN(shadow/cap)) * cap"
        .Font.Color = RGB(255, 99, 71)
        .Font.Bold = True
        .Font.Size = 10
    End With

    With wsD.Cells(painStartRow, segments.Count + 3)
        .Value = "PAIN CAP%"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    With wsD.Cells(painStartRow, segments.Count + 4)
        .Value = PAIN_CAP
        .Font.Color = RGB(0, 191, 255)
        .Font.Bold = True
    End With

    With wsD.Cells(painHeaderRow, 1)
        .Value = "Day"
        .Font.Color = RGB(150, 150, 150)
    End With
    For d = 1 To maxDays
        With wsD.Cells(painHeaderRow + d, 1)
            .Value = d
            .Font.Color = RGB(200, 200, 200)
        End With
    Next d

    For s = 1 To segments.Count
        colIdx = s + 1
        With wsD.Cells(painHeaderRow, colIdx)
            .Value = segDates(s)
            .Font.Color = RGB(180, 180, 180)
            .Font.Size = 8
        End With
    Next s

    ' --- Compute pain values ---
    For s = 1 To segments.Count
        arr = segments(s)
        segLenInner = UBound(arr) + 1
        colIdx = s + 1
        cumRet = 0

        For d = 0 To segLenInner - 1
            xVal = d + 1
            cumRet = cumRet + arr(d)
            ' BUG FIX: previous version had '_' line continuation but missing '+'.
            ' Now the '+' is explicit so e^sqrt(x) is properly added.
            shadowVal = (Exp(xVal) / (xVal ^ 3 + WorksheetFunction.Pi() ^ 3) + _
                         Exp(Sqr(xVal))) * cumRet

            If Abs(shadowVal) >= PAIN_CAP Then
                painVal = Sgn(shadowVal) * PAIN_OVERFLOW
            Else
                painVal = Tan(WorksheetFunction.Asin(shadowVal / PAIN_CAP)) * PAIN_CAP
            End If

            painDataRow = painHeaderRow + xVal
            With wsD.Cells(painDataRow, colIdx)
                .Value = Round(painVal, 4)
                .NumberFormat = "0.0000"
                .Font.Size = 9
                depth = Abs(painVal)
                If depth >= PAIN_OVERFLOW Then
                    .Font.Color = RGB(255, 0, 0): .Font.Bold = True
                ElseIf depth > 5 Then
                    .Font.Color = RGB(255, 40, 20)
                ElseIf depth > 2 Then
                    .Font.Color = RGB(240, 60, 30)
                ElseIf depth > 0.5 Then
                    .Font.Color = RGB(210, 80, 50)
                Else
                    .Font.Color = RGB(180, 100, 80)
                End If
            End With

            If painVal < minPain Then minPain = painVal
        Next d
    Next s

    ' --- Pain stats ---
    Dim painStatCol As Long: painStatCol = segments.Count + 3
    With wsD.Cells(painHeaderRow + 1, painStatCol)
        .Value = "PAIN STATS"
        .Font.Color = RGB(255, 99, 71)
        .Font.Bold = True
    End With
    With wsD.Cells(painHeaderRow + 2, painStatCol)
        .Value = "Max Pain"
        .Font.Color = RGB(150, 150, 150)
    End With
    With wsD.Cells(painHeaderRow + 2, painStatCol + 1)
        .Value = Round(minPain, 4)
        .NumberFormat = "0.0000"
        .Font.Color = RGB(255, 69, 0)
    End With

    If minCum <> 0 Then
        With wsD.Cells(painHeaderRow + 3, painStatCol)
            .Value = "Pain/DD"
            .Font.Color = RGB(150, 150, 150)
        End With
        With wsD.Cells(painHeaderRow + 3, painStatCol + 1)
            .Value = Format(minPain / minCum, "0.00") & "x"
            .Font.Color = RGB(255, 69, 0)
        End With
    End If

    ' --- Summary panel ---
    Dim sumRow As Long: sumRow = DATA_ROW - 1
    Dim sumCol As Long: sumCol = segments.Count + 3
    With wsD.Cells(sumRow, sumCol)
        .Value = "DRAWDOWN SUMMARY v2"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    wsD.Cells(sumRow + 1, sumCol).Value = "Episodes"
    wsD.Cells(sumRow + 1, sumCol + 1).Value = segments.Count
    wsD.Cells(sumRow + 2, sumCol).Value = "Max Days"
    wsD.Cells(sumRow + 2, sumCol + 1).Value = maxDays
    wsD.Cells(sumRow + 3, sumCol).Value = "Min f(x)"
    wsD.Cells(sumRow + 3, sumCol + 1).Value = Format(minCum, "0.0000")
    wsD.Cells(sumRow + 4, sumCol).Value = "Rule"
    wsD.Cells(sumRow + 4, sumCol + 1).Value = WIN_MIN_NEG & "/" & WIN_SIZE & " neg"

    wsD.Cells(sumRow + 6, sumCol).Value = "Start"
    wsD.Cells(sumRow + 6, sumCol + 1).Value = "Days"
    wsD.Cells(sumRow + 6, sumCol + 2).Value = "Sum Neg%"
    wsD.Cells(sumRow + 6, sumCol).Font.Color = RGB(255, 192, 0)
    wsD.Cells(sumRow + 6, sumCol + 1).Font.Color = RGB(255, 192, 0)
    wsD.Cells(sumRow + 6, sumCol + 2).Font.Color = RGB(255, 192, 0)

    Dim sumNeg As Double
    For s = 1 To segments.Count
        arr = segments(s)
        sumNeg = 0
        For k = 0 To UBound(arr)
            sumNeg = sumNeg + arr(k)   ' arr already zero-filled for positives
        Next k
        wsD.Cells(sumRow + 6 + s, sumCol).Value = segDates(s)
        wsD.Cells(sumRow + 6 + s, sumCol + 1).Value = UBound(arr) + 1
        wsD.Cells(sumRow + 6 + s, sumCol + 2).Value = Format(sumNeg * 100, "0.00") & "%"
        wsD.Cells(sumRow + 6 + s, sumCol + 2).Font.Color = RGB(200, 80, 80)
    Next s

    For ci = sumCol To sumCol + 2
        wsD.Columns(ci).ColumnWidth = 14
    Next ci

CleanExit:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If segments.Count > 0 Then
        Application.StatusBar = "Drawdown Shadow v2 done: " & segments.Count & " episodes"
        MsgBox "Done! Found " & segments.Count & " drawdown segments." & vbCrLf & _
               "Rule: 5-day window with >= 4 negative days" & vbCrLf & _
               "Positive returns within segment are ignored in cumRet.", vbInformation
    End If
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
