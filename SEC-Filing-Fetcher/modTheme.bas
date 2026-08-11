Attribute VB_Name = "modTheme"
Option Explicit

Private Const THEME_FONT As String = "Yu Gothic"
Private Const THEME_BLACK As Long = 0                 ' RGB(0,0,0)
Private Const THEME_WHITE As Long = 16777215           ' RGB(255,255,255)
Private Const THEME_ORANGE As Long = 42495             ' RGB(255,165,0)
Private Const THEME_GRIDLINE As Long = 2302755         ' RGB(35,35,35) -- deliberately faint against the black plot area
Private Const THEME_AXISLINE As Long = 7895160         ' RGB(120,120,120)
Private Const THEME_BORDER As Long = 5263440           ' RGB(80,80,80) -- dark gray cell grid

' Cycles through a small set of orange shades so neighboring metric charts are
' visually distinguishable at a glance while staying within the requested
' black/orange palette.
Public Function OrangePalette(ByVal idx As Long) As Long
    Dim colors As Variant
    colors = Array(RGB(255, 165, 0), RGB(255, 140, 0), RGB(230, 126, 34), RGB(255, 127, 80), _
                    RGB(211, 84, 0), RGB(243, 156, 18), RGB(255, 99, 71), RGB(184, 89, 18))
    OrangePalette = colors(idx Mod 8)
End Function

' Black background / white body text / orange bold header row, Calibri throughout.
' Also blackens a generous margin well beyond the actual data so scrolling or a
' wider window doesn't hit an abrupt white cutoff outside the used range.
Public Sub ApplyDarkTheme(ByVal ws As Worksheet, ByVal dataRange As Range, Optional ByVal headerRow As Long = 1)
    With dataRange
        .Interior.Color = THEME_BLACK
        .Font.Color = THEME_WHITE
        .Font.Name = THEME_FONT
    End With
    Call ApplyGridBorder(dataRange)

    Dim lastDataRow As Long
    lastDataRow = dataRange.Row - 1 + dataRange.Rows.Count
    Dim extendRows As Long
    extendRows = lastDataRow + 200
    If extendRows < 500 Then extendRows = 500

    ws.Range(ws.Cells(1, 1), ws.Cells(extendRows, 60)).Interior.Color = THEME_BLACK

    Dim hdrRange As Range
    Set hdrRange = ws.Range(ws.Cells(headerRow, dataRange.Column), ws.Cells(headerRow, dataRange.Column + dataRange.Columns.Count - 1))
    hdrRange.Font.Color = THEME_ORANGE
    hdrRange.Font.Bold = True
End Sub

' Thin dark-gray grid lines around every cell in range, visible against the
' black fill (Excel's default gridline color is invisible on black).
Public Sub ApplyGridBorder(ByVal r As Range)
    With r.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = THEME_BORDER
    End With
End Sub

' seriesColors: array of RGB Long values, one per series (cycles if there are
' more series than colors, e.g. the 2-series price chart).
Public Sub ApplyDarkThemeToChart(ByVal ch As Chart, ByVal seriesColors As Variant)
    ch.ChartArea.Interior.Color = THEME_BLACK
    ch.ChartArea.Font.Name = THEME_FONT
    ch.ChartArea.Font.Color = THEME_WHITE

    On Error Resume Next
    ch.PlotArea.Interior.Color = THEME_BLACK
    On Error GoTo 0

    If ch.HasTitle Then
        ch.ChartTitle.Font.Color = THEME_ORANGE
        ch.ChartTitle.Font.Name = THEME_FONT
        ch.ChartTitle.Font.Size = 10
        ch.ChartTitle.Font.Bold = False
    End If

    On Error Resume Next
    Dim axC As Axis, axV As Axis
    Set axC = ch.Axes(xlCategory)
    axC.TickLabels.Font.Color = THEME_WHITE
    axC.TickLabels.Font.Name = THEME_FONT
    axC.TickLabels.Font.Size = 8
    axC.Line.ForeColor.RGB = THEME_AXISLINE

    Set axV = ch.Axes(xlValue)
    axV.TickLabels.Font.Color = THEME_WHITE
    axV.TickLabels.Font.Name = THEME_FONT
    axV.TickLabels.Font.Size = 8
    axV.Line.ForeColor.RGB = THEME_AXISLINE
    axV.MajorGridlines.Border.Color = THEME_GRIDLINE

    ' Excel's auto-scaled gridline spacing tends to be needlessly fine (7-9
    ' lines); force a coarser, simpler ~4-line grid instead. MinimumScale/
    ' MaximumScale read out Excel's own auto-computed bounds at this point,
    ' so this works generically across every metric's very different scales.
    Dim axRange As Double
    axRange = axV.MaximumScale - axV.MinimumScale
    If axRange > 0 Then axV.MajorUnit = axRange / 4
    On Error GoTo 0

    If ch.HasLegend Then
        ch.Legend.Font.Color = THEME_WHITE
        ch.Legend.Font.Name = THEME_FONT
    End If

    Dim lo As Long, hi As Long, spread As Long
    lo = LBound(seriesColors)
    hi = UBound(seriesColors)
    spread = hi - lo + 1

    Dim i As Long
    For i = 1 To ch.SeriesCollection.Count
        Dim colorForThis As Long
        colorForThis = seriesColors(lo + ((i - 1) Mod spread))
        Call ColorSeries(ch.SeriesCollection(i), colorForThis)
    Next i
End Sub

' Rounds a raw axis step up to the nearest 1 / 2 / 2.5 / 5 x 10^k, so a
' data-fitted axis still gets round tick labels (180M, 200M, 220M...) instead
' of the arbitrary ones a raw span/4 would produce (182.74M, 206.39M...).
Private Function NiceStep(ByVal raw As Double) As Double
    NiceStep = 0
    If raw <= 0 Then Exit Function

    Dim mag As Double, norm As Double
    mag = 10 ^ Int(Log(raw) / Log(10#))
    norm = raw / mag

    Dim mult As Double
    If norm <= 1 Then
        mult = 1
    ElseIf norm <= 2 Then
        mult = 2
    ElseIf norm <= 2.5 Then
        mult = 2.5
    ElseIf norm <= 5 Then
        mult = 5
    Else
        mult = 10
    End If

    NiceStep = mult * mag
End Function

' Fits a value axis to the data actually plotted on it, instead of leaving it to
' Excel's auto-scale -- which anchors nearly every axis at zero. Measured on a
' real AXTI fetch, that left 29 of the Dashboard's 47 charts using under 60% of
' the plot height (股東權益 179M-275M drawn on a 0-300M axis reads as a flat
' line even though the underlying value moved 53%).
'
' Pads by 5% of the data span at each end so extremes don't sit on the frame,
' then snaps both ends outward to a round step. Returns the resulting number of
' major units, so a caller with two value axes can re-fit the other one to the
' same count via forceUnits and get gridlines that coincide.
'
' Leaves the axis untouched (returns 0) when the series is flat -- a company
' that pays no dividend plots a straight line at zero, which has no span to fit
' and would ask Excel for Min = Max.
Public Function FitValueAxis(ByVal ax As Axis, ByVal lo As Double, ByVal hi As Double, Optional ByVal forceUnits As Long = 0) As Long
    FitValueAxis = 0

    Dim span As Double
    span = hi - lo
    If span <= 0 Then Exit Function

    Dim pad As Double
    pad = span * 0.05

    Dim padLo As Double, padHi As Double
    padLo = lo - pad
    padHi = hi + pad
    ' Don't let padding drag an all-positive series below zero -- a stock price
    ' axis that starts at -5 is worse than losing that 5% of headroom.
    If lo >= 0 And padLo < 0 Then padLo = 0

    Dim padSpan As Double
    padSpan = padHi - padLo
    If padSpan <= 0 Then Exit Function

    ' Snapping the ends out to a round step costs some of the fit, and how much
    ' depends entirely on where the data happens to fall relative to that step.
    ' Try 4, 5 and 6 divisions and keep whichever wastes least: on real numbers
    ' 流動比率 2.09-4.21 snaps to 1..5 at four divisions (53% used) but to
    ' 1.5..4.5 at five (71%). Ties keep the coarsest grid, and six is still
    ' fewer gridlines than the 7-9 Excel picks on its own.
    Dim bestTick As Double, bestLo As Double, bestHi As Double, bestUse As Double
    bestUse = -1

    Dim target As Long
    For target = 4 To 6
        Dim tick As Double
        tick = NiceStep(padSpan / target)
        If tick > 0 Then
            Dim tryLo As Double, tryHi As Double
            tryLo = Int(padLo / tick) * tick
            tryHi = -Int(-padHi / tick) * tick
            If lo >= 0 And tryLo < 0 Then tryLo = 0
            If tryHi > tryLo Then
                Dim used As Double
                used = span / (tryHi - tryLo)
                If used > bestUse + 0.0001 Then
                    bestUse = used
                    bestTick = tick
                    bestLo = tryLo
                    bestHi = tryHi
                End If
            End If
        End If
    Next target

    If bestUse < 0 Then Exit Function

    Dim units As Long
    units = CLng((bestHi - bestLo) / bestTick)
    If units < 1 Then units = 1
    If forceUnits > units Then
        bestHi = bestLo + bestTick * forceUnits
        units = forceUnits
    End If

    ax.MinimumScale = bestLo
    ax.MaximumScale = bestHi
    ax.MajorUnit = bestTick
    FitValueAxis = units
End Function

' Tries both Fill (area/bar charts) and Line (line charts) forecolor -- exactly
' one of these applies per chart type, and the other errors harmlessly.
Private Sub ColorSeries(ByVal ser As Series, ByVal color As Long)
    On Error Resume Next
    ser.Format.Fill.ForeColor.RGB = color
    ser.Format.Line.ForeColor.RGB = color
    On Error GoTo 0
End Sub
