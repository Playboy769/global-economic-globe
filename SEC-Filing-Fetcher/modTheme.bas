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

' Tries both Fill (area/bar charts) and Line (line charts) forecolor -- exactly
' one of these applies per chart type, and the other errors harmlessly.
Private Sub ColorSeries(ByVal ser As Series, ByVal color As Long)
    On Error Resume Next
    ser.Format.Fill.ForeColor.RGB = color
    ser.Format.Line.ForeColor.RGB = color
    On Error GoTo 0
End Sub
