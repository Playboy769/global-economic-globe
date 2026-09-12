Attribute VB_Name = "modvolatility"
Option Explicit

'=============================================================================
'  TICKERS VOLATILITY (nav code VT) - sheet "Tickers Volatility"
'-----------------------------------------------------------------------------
'  v2 (2026-09-12): rebuilt in the RR4 style. What changed:
'    * the three lookbacks (100 / 200 / 1000 trading days) are now three
'      COLUMNS of one table instead of three stacked blocks, so short-term
'      against long-term is one glance instead of three scrolls
'    * Skewness / Kurtosis / Max Up Streak / Max Down Streak dropped - daily
'      higher moments flip sign with the sample and streaks are not actionable
'    * new 20D VOL PERCENTILE: where the current 20-day vol sits inside that
'      same column's window - the one number that answers "is it high NOW"
'    * the 12 charts used to be hand-drawn objects pointing at fixed cells.
'      They are rebuilt from code on every run now (DrawVtCharts) - the old
'      set had drifted: the 200-day Max/Min chart was plotting B16:B17, i.e.
'      the 100-day numbers, and the three Risk charts fed 4 values into a
'      6-label category axis.
'
'  Page layout (drawn from row 1; modNav adds the 3 bar rows afterwards -
'  the sheet code does NavStrip before this runs and NavAdd after):
'    row 1        title
'    row 2        A "TICKER <GO>"   B input cell
'    row 4        header   A STATISTIC | B 100D | C 200D | D 1000D
'    rows 5-14    the statistics (see VtStatLabels - the order is load
'                 bearing, each chart takes a contiguous slice of it)
'    row 16/17    RETURN DISTRIBUTION title / header
'    rows 18-39   22 bins, one shared RANGE column + one Freq% per period
'    row 57+      formula notes (below the chart band - column A text would
'                 otherwise run underneath the charts and be unreadable)
'    column F on  chart grid, one band per period, 4 charts across
'=============================================================================

Private Const VT_SHEET      As String = "Tickers Volatility"
Private Const VT_TITLE_ROW  As Long = 1
Private Const VT_INPUT_ROW  As Long = 2
Private Const VT_HDR_ROW    As Long = 4
Private Const VT_STAT_FIRST As Long = 5
Private Const VT_STAT_N     As Long = 10
Private Const VT_DIST_TITLE As Long = 16
Private Const VT_DIST_HDR   As Long = 17
Private Const VT_DIST_FIRST As Long = 18
Private Const VT_BINS       As Long = 22
Private Const VT_NOTE_ROW   As Long = 58     ' below the chart band, else the charts cover the text
Private Const VT_ROLL       As Long = 20
Private Const VT_CHART_PFX  As String = "VT_"

' Rows inside the statistics block that each chart plots, as offsets from
' VT_STAT_FIRST. Keep these in step with VtStatLabels - the whole point of
' the fixed order is that every chart is one contiguous slice.
Private Const VT_VOL_FROM   As Long = 0     ' mean / std dev / annualized
Private Const VT_VOL_TO     As Long = 2
Private Const VT_RISK_FROM  As Long = 3     ' downside dev / MDD / cur 20D / max 20D
Private Const VT_RISK_TO    As Long = 6
Private Const VT_MM_FROM    As Long = 7     ' max / min daily change
Private Const VT_MM_TO      As Long = 8

Private Const VT_TXT        As Long = 14540253    ' RGB(221,221,221)
Private Const VT_DIM        As Long = 9211020     ' RGB(140,140,140)
Private Const VT_UP         As Long = 3947980     ' RGB(204,60,60)  - TW: up is red
Private Const VT_DOWN       As Long = 7126016     ' RGB(0,190,108)  - TW: down is green

Private Function VtLookbacks() As Variant
    VtLookbacks = Array(100, 200, 1000)
End Function

Private Function VtColLabels() As Variant
    VtColLabels = Array("100D", "200D", "1000D")
End Function

' One colour per period - same hue, three depths, so the three bands read as
' one family instead of the old orange / pink / red mix.
Private Function VtTone(ByVal k As Long) As Long
    Select Case k
        Case 0:    VtTone = RGB(235, 130, 30)
        Case 1:    VtTone = RR4_ACCENT
        Case Else: VtTone = RGB(175, 92, 10)
    End Select
End Function

Private Function VtStatLabels() As Variant
    VtStatLabels = Array( _
        "MEAN (DAILY)", "STD DEV (DAILY)", "ANNUALIZED VOLATILITY", _
        "DOWNSIDE DEVIATION", "MAX DRAWDOWN (MDD)", _
        "CURRENT 20D VOLATILITY", "MAX 20D VOLATILITY", _
        "MAX DAILY CHG%", "MIN DAILY CHG%", "20D VOL PERCENTILE")
End Function

' The typed input cell. This module always runs with the nav bar stripped,
' so it is the page address; the sheet code offsets it with NavOffset.
Private Function VtInputAddr() As String
    VtInputAddr = "B" & VT_INPUT_ROW
End Function

'=============================================================================
Sub UpdateVolatilityAnalysis_Pro()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim allPrices() As Double, allDates() As Date
    Dim lb As Variant, k As Long

    Set ws = ThisWorkbook.Sheets(VT_SHEET)
    Ticker = UCase$(Trim$(CStr(ws.Range(VtInputAddr()).Value)))
    If Ticker = "" Then Call NavNotify("VT: type a ticker in the input cell", True): Exit Sub

    If Not GetHistoricalData(Ticker, allDates, allPrices) Then
        Call NavNotify("VT: could not download price history for " & Ticker, True)
        Exit Sub
    End If
    If UBound(allPrices) + 1 < 25 Then
        Call NavNotify("VT: not enough price history for " & Ticker & " (need 25+ days)", True)
        Exit Sub
    End If

    ' ResetVtSheet writes the ticker back into the input cell, which would
    ' fire the sheet's Worksheet_Change and re-enter this routine. The sheet
    ' code already turns events off, but a run started any other way must not
    ' depend on that - see the same trap on the CC page.
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fin

    Call ResetVtSheet(ws, Ticker)

    lb = VtLookbacks()
    For k = LBound(lb) To UBound(lb)
        Call AnalyzeAndOutput(ws, allPrices, CLng(lb(k)), 2 + k)
    Next k

    Call DrawVtNotes(ws)
    Call DrawVtCharts(ws)

Fin:
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    If Err.Number <> 0 Then
        Call NavNotify("VT error " & Err.Number & ": " & Err.Description, True)
    Else
        Call NavNotify("VT done - 100 / 200 / 1000-day volatility for " & Ticker)
    End If
End Sub

'=============================================================================
'  Clear the page and lay down everything that does not depend on the data:
'  base style, title, input cell, both header rows, the label column.
'  The old version only did ClearContents on a hard-coded A4:F100, so stale
'  formats and anything written past row 100 survived a rebuild.
'=============================================================================
Private Sub ResetVtSheet(ws As Worksheet, ByVal Ticker As String)
    Dim i As Long, r As Long
    Dim cols As Variant: cols = VtColLabels()

    Call DeleteVtCharts(ws)
    ws.Cells.Clear
    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = VT_TXT
        .Font.Name = "Consolas"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlRight
    End With
    For r = 1 To VT_NOTE_ROW + 10
        ws.Rows(r).RowHeight = 17
    Next r
    ws.Rows(VT_TITLE_ROW).RowHeight = 24
    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B:D").ColumnWidth = 13
    ws.Columns("E").ColumnWidth = 6

    ' ---- title ----
    With ws.Cells(VT_TITLE_ROW, 1)
        .Value = "TICKERS VOLATILITY"
        .HorizontalAlignment = xlLeft
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RR4_ACCENT
    End With

    ' ---- input ----
    With ws.Cells(VT_INPUT_ROW, 1)
        .Value = "TICKER <GO>"
        .HorizontalAlignment = xlLeft
        .Font.Color = VT_DIM
    End With
    With ws.Cells(VT_INPUT_ROW, 2)
        .Value = Ticker
        .HorizontalAlignment = xlLeft
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
    End With
    Call VtLine(ws, VT_INPUT_ROW, 1, 4)

    ' ---- statistics header + labels ----
    Call VtHeader(ws, VT_HDR_ROW, "STATISTIC", cols)
    Dim labs As Variant: labs = VtStatLabels()
    For i = 0 To VT_STAT_N - 1
        With ws.Cells(VT_STAT_FIRST + i, 1)
            .Value = labs(i)
            .HorizontalAlignment = xlLeft
        End With
        If i Mod 2 = 1 Then
            ws.Range(ws.Cells(VT_STAT_FIRST + i, 1), ws.Cells(VT_STAT_FIRST + i, 4)) _
              .Interior.Color = RGB(14, 14, 14)
        End If
    Next i
    Call VtLine(ws, VT_STAT_FIRST + VT_STAT_N - 1, 1, 4)

    ' ---- distribution header + bin labels ----
    With ws.Cells(VT_DIST_TITLE, 1)
        .Value = "RETURN DISTRIBUTION (FREQ %)"
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.Color = RR4_ACCENT
    End With
    Call VtHeader(ws, VT_DIST_HDR, "RANGE", cols)
    For i = 0 To VT_BINS - 1
        With ws.Cells(VT_DIST_FIRST + i, 1)
            .Value = "'" & VtBinLabel(i)
            .HorizontalAlignment = xlLeft
        End With
        If i Mod 2 = 1 Then
            ws.Range(ws.Cells(VT_DIST_FIRST + i, 1), ws.Cells(VT_DIST_FIRST + i, 4)) _
              .Interior.Color = RGB(14, 14, 14)
        End If
    Next i
    Call VtLine(ws, VT_DIST_FIRST + VT_BINS - 1, 1, 4)
End Sub

Private Sub VtHeader(ws As Worksheet, ByVal r As Long, ByVal firstLabel As String, ByVal cols As Variant)
    Dim c As Long
    With ws.Cells(r, 1)
        .Value = firstLabel
        .HorizontalAlignment = xlLeft
    End With
    For c = 0 To UBound(cols)
        ws.Cells(r, 2 + c).Value = cols(c)
    Next c
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Font.Bold = True
        .Font.Color = RR4_ACCENT
    End With
    Call VtLine(ws, r, 1, 4)
End Sub

Private Sub VtLine(ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RR4_LINE
    End With
End Sub

Private Function VtBinLabel(ByVal i As Long) As String
    Dim lower As Double
    If i = 0 Then
        VtBinLabel = "< -10%"
    ElseIf i = VT_BINS - 1 Then
        VtBinLabel = "> 10%"
    Else
        lower = -0.1 + (i - 1) * 0.01
        VtBinLabel = Format(lower, "0%") & " ~ " & Format(lower + 0.01, "0%")
    End If
End Function

'=============================================================================
'  One lookback window -> one column of the statistics table and one column
'  of the distribution table.
'=============================================================================
Private Sub AnalyzeAndOutput(ws As Worksheet, ByRef allPrices() As Double, _
                             ByVal lookbackDays As Long, ByVal outCol As Long)
    Dim prices() As Double, returns() As Double
    Dim i As Long, j As Long, n As Long, totalData As Long, useDays As Long

    totalData = UBound(allPrices) + 1
    useDays = lookbackDays
    If totalData < useDays Then useDays = totalData
    If useDays < 3 Then Exit Sub

    ReDim prices(0 To useDays - 1)
    For i = 0 To useDays - 1
        prices(i) = allPrices(UBound(allPrices) - useDays + 1 + i)
    Next i

    n = UBound(prices)
    ReDim returns(0 To n - 1)
    For i = 1 To n
        If prices(i - 1) <> 0 Then
            returns(i - 1) = (prices(i) - prices(i - 1)) / prices(i - 1)
        Else
            returns(i - 1) = 0
        End If
    Next i

    ' ---- max drawdown ----
    Dim peak As Double, dd As Double, mdd As Double
    peak = prices(0): mdd = 0
    For i = 1 To n
        If prices(i) > peak Then peak = prices(i)
        If peak <> 0 Then
            dd = (peak - prices(i)) / peak
            If dd > mdd Then mdd = dd
        End If
    Next i

    ' ---- downside deviation ----
    Dim downReturns() As Double, downCount As Long
    downCount = 0
    For i = LBound(returns) To UBound(returns)
        If returns(i) < 0 Then
            ReDim Preserve downReturns(0 To downCount)
            downReturns(downCount) = returns(i)
            downCount = downCount + 1
        End If
    Next i

    ' ---- rolling 20-day annualized vol: current, max, and where current
    '      sits inside this column's own window (the percentile) ----
    Dim rollWindow(1 To VT_ROLL) As Double
    Dim rolls() As Double
    Dim rollVol As Double, currRollVol As Double, maxRollVol As Double
    Dim rollCount As Long, atOrBelow As Long
    maxRollVol = 0: rollCount = 0: atOrBelow = 0
    If UBound(returns) >= VT_ROLL - 1 Then
        ReDim rolls(0 To UBound(returns) - VT_ROLL + 1)
        For i = VT_ROLL - 1 To UBound(returns)
            For j = 1 To VT_ROLL
                rollWindow(j) = returns(i - VT_ROLL + j)
            Next j
            rollVol = Application.WorksheetFunction.StDev_S(rollWindow) * Sqr(252)
            rolls(rollCount) = rollVol
            rollCount = rollCount + 1
            If rollVol > maxRollVol Then maxRollVol = rollVol
            If i = UBound(returns) Then currRollVol = rollVol
        Next i
        For i = 0 To rollCount - 1
            If rolls(i) <= currRollVol Then atOrBelow = atOrBelow + 1
        Next i
    End If

    Dim pctile As Double
    If rollCount > 0 Then pctile = atOrBelow / rollCount

    ' ---- basic moments ----
    Dim mean As Double, stdDev As Double, annVol As Double
    Dim maxRet As Double, minRet As Double, downDev As Double
    On Error Resume Next
    mean = Application.WorksheetFunction.Average(returns)
    stdDev = Application.WorksheetFunction.StDev_S(returns)
    maxRet = Application.WorksheetFunction.Max(returns)
    minRet = Application.WorksheetFunction.Min(returns)
    If downCount > 1 Then downDev = Application.WorksheetFunction.StDev_S(downReturns)
    On Error GoTo 0
    annVol = stdDev * Sqr(252)

    ' ---- statistics column ----
    Dim vals As Variant
    vals = Array(mean, stdDev, annVol, downDev, -mdd, currRollVol, maxRollVol, _
                 maxRet, minRet, pctile)
    For i = 0 To VT_STAT_N - 1
        With ws.Cells(VT_STAT_FIRST + i, outCol)
            .Value = vals(i)
            .NumberFormat = "0.00%"
        End With
    Next i
    ' a percentile is a rank, not a return - one decimal is plenty
    ws.Cells(VT_STAT_FIRST + VT_STAT_N - 1, outCol).NumberFormat = "0.0%"
    ' a drawdown is always down, the best day always up: colour them TW-style
    ws.Cells(VT_STAT_FIRST + 4, outCol).Font.Color = VT_DOWN
    ws.Cells(VT_STAT_FIRST + 7, outCol).Font.Color = VT_UP
    ws.Cells(VT_STAT_FIRST + 8, outCol).Font.Color = VT_DOWN
    ' current vol above this window's own 80th percentile is worth noticing
    If pctile >= 0.8 Then
        With ws.Cells(VT_STAT_FIRST + VT_STAT_N - 1, outCol)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
        End With
    End If

    ' ---- distribution column ----
    Dim bins() As Long, retVal As Double, binIdx As Long, total As Long
    ReDim bins(0 To VT_BINS - 1)
    total = UBound(returns) + 1
    For i = LBound(returns) To UBound(returns)
        retVal = returns(i)
        If retVal < -0.1 Then
            bins(0) = bins(0) + 1
        ElseIf retVal >= 0.1 Then
            bins(VT_BINS - 1) = bins(VT_BINS - 1) + 1
        Else
            binIdx = Int((retVal + 0.1) / 0.01) + 1
            If binIdx >= 1 And binIdx <= VT_BINS - 2 Then bins(binIdx) = bins(binIdx) + 1
        End If
    Next i
    For i = 0 To VT_BINS - 1
        With ws.Cells(VT_DIST_FIRST + i, outCol)
            If total > 0 Then
                .Value = bins(i) / total
            Else
                .Value = 0
            End If
            .NumberFormat = "0.00%"
        End With
    Next i

    ' the window actually used, in case the history was shorter than asked for
    On Error Resume Next
    ws.Cells(VT_HDR_ROW, outCol).ClearComments
    ws.Cells(VT_HDR_ROW, outCol).AddComment Text:="days used: " & useDays
    ws.Cells(VT_HDR_ROW, outCol).Comment.Visible = False
    On Error GoTo 0
End Sub

'=============================================================================
'  Formula notes, same habit as the V / C / CC pages.
'=============================================================================
Private Sub DrawVtNotes(ws As Worksheet)
    Dim notes As Variant, i As Long
    notes = Array( _
        "ANNUALIZED VOLATILITY = daily std dev x sqrt(252)   - the headline number; read the three columns together to see if the stock is calmer or wilder than its own history", _
        "DOWNSIDE DEVIATION    = std dev of the negative days only   - half the volatility of something that only jumps up is not risk", _
        "MAX DRAWDOWN (MDD)    = deepest peak-to-trough fall inside the window   - the hole a position has to be able to sit through", _
        "CURRENT / MAX 20D VOL = annualized std dev of the last 20 days, and the highest that rolling figure reached inside the window", _
        "20D VOL PERCENTILE    = share of the window's rolling 20-day vols at or below the current one   - 90% means only 1 stretch in 10 was this jumpy; lit orange from 80%", _
        "RETURN DISTRIBUTION   = share of days landing in each 1% bucket, both tails bucketed at -10% / +10%   - fat ends mean the std dev understates the real risk")
    With ws.Cells(VT_NOTE_ROW - 1, 1)
        .Value = "HOW THESE ARE CALCULATED"
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.Color = RR4_ACCENT
    End With
    For i = LBound(notes) To UBound(notes)
        With ws.Cells(VT_NOTE_ROW + i, 1)
            .Value = notes(i)
            .HorizontalAlignment = xlLeft
            .Font.Color = VT_DIM
            .Font.Size = 8
        End With
    Next i
End Sub

'=============================================================================
'  Charts. Four per period, three periods, all rebuilt from scratch so their
'  source ranges can never drift away from the table again.
'=============================================================================
Private Sub DeleteVtCharts(ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.ChartObjects().count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Sub DrawVtCharts(ws As Worksheet)
    Dim k As Long, baseLeft As Double, topPt As Double
    Dim cols As Variant: cols = VtColLabels()
    Const BAND_H As Double = 296
    Const BAND_GAP As Double = 14
    Const W_HIST As Double = 640
    Const W_SMALL As Double = 300
    Const H_GAP As Double = 12

    baseLeft = ws.Cells(1, 6).Left        ' column F
    For k = 0 To 2
        topPt = 6 + k * (BAND_H + BAND_GAP)
        Call VtChart(ws, "HIST_" & k, "CHG% DISTRIBUTION - " & cols(k), k, _
             baseLeft, topPt, W_HIST, BAND_H, _
             VT_DIST_FIRST, VT_DIST_FIRST + VT_BINS - 1, 2 + k, True, False)
        Call VtChart(ws, "VOL_" & k, "VOLATILITY - " & cols(k), k, _
             baseLeft + W_HIST + H_GAP, topPt, W_SMALL, BAND_H, _
             VT_STAT_FIRST + VT_VOL_FROM, VT_STAT_FIRST + VT_VOL_TO, 2 + k, False, False)
        Call VtChart(ws, "RISK_" & k, "RISK ASSESSMENT - " & cols(k), k, _
             baseLeft + W_HIST + W_SMALL + 2 * H_GAP, topPt, W_SMALL, BAND_H, _
             VT_STAT_FIRST + VT_RISK_FROM, VT_STAT_FIRST + VT_RISK_TO, 2 + k, False, False)
        Call VtChart(ws, "MAXMIN_" & k, "MAX / MIN DAILY - " & cols(k), k, _
             baseLeft + W_HIST + 2 * W_SMALL + 3 * H_GAP, topPt, W_SMALL, BAND_H, _
             VT_STAT_FIRST + VT_MM_FROM, VT_STAT_FIRST + VT_MM_TO, 2 + k, False, True)
    Next k
End Sub

'  rowFrom/rowTo  the contiguous slice of the table this chart plots
'  valCol         which period column supplies the values
'  dense          True for the 22-bin histogram (smaller labels, tight bars)
'  upDown         True to colour the first point red and the second green
Private Sub VtChart(ws As Worksheet, ByVal nameSuffix As String, ByVal title As String, _
                    ByVal tone As Long, ByVal L As Double, ByVal T As Double, _
                    ByVal W As Double, ByVal H As Double, _
                    ByVal rowFrom As Long, ByVal rowTo As Long, ByVal valCol As Long, _
                    ByVal dense As Boolean, ByVal upDown As Boolean)
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(L, T, W, H)
    co.Name = VT_CHART_PFX & nameSuffix
    co.Placement = xlMove              ' NavAdd inserts rows above - move, never resize

    With co.Chart
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        .ChartType = xlColumnClustered
        With .SeriesCollection.NewSeries
            .XValues = ws.Range(ws.Cells(rowFrom, 1), ws.Cells(rowTo, 1))
            .Values = ws.Range(ws.Cells(rowFrom, valCol), ws.Cells(rowTo, valCol))
        End With

        .HasLegend = False
        .HasTitle = True
        With .ChartTitle
            .Text = title
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RR4_ACCENT
        End With
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .ChartArea.Format.Line.ForeColor.RGB = RR4_LINE
        .ChartArea.Font.Name = "Consolas"
        .ChartArea.Font.Color = VT_TXT
        .PlotArea.Format.Fill.Visible = msoFalse

        With .Axes(xlCategory)
            .TickLabels.Font.Size = IIf(dense, 7, 8)
            .TickLabels.Font.Color = VT_DIM
            .Format.Line.ForeColor.RGB = RR4_LINE
            ' MDD and the min-day bar go below zero; left at the default the
            ' category labels sit on the zero line and collide with the bars
            .TickLabelPosition = xlTickLabelPositionLow
        End With
        With .Axes(xlValue)
            .TickLabels.Font.Size = 8
            .TickLabels.Font.Color = VT_DIM
            .TickLabels.NumberFormat = "0%"
            .Format.Line.ForeColor.RGB = RR4_LINE
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(35, 35, 35)
        End With

        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = VtTone(tone)
            .Format.Line.Visible = msoFalse
            If Not dense Then
                .HasDataLabels = True
                With .DataLabels
                    .NumberFormat = "0.00%"
                    .Font.Size = 8
                    .Font.Color = VT_TXT
                End With
            End If
            ' TW convention, same as the RR4 weight bars: up red, down green
            If upDown Then
                .Points(1).Format.Fill.ForeColor.RGB = VT_UP
                .Points(2).Format.Fill.ForeColor.RGB = VT_DOWN
            End If
        End With
        If dense Then
            .ChartGroups(1).GapWidth = 20
        ElseIf rowTo - rowFrom <= 1 Then
            .ChartGroups(1).GapWidth = 220      ' two bars at gap 90 read as slabs
        Else
            .ChartGroups(1).GapWidth = 110
        End If
    End With
End Sub

'=============================================================================
'  Price history from Yahoo. A bare TW number is tried as .TW first and then
'  as .TWO - the old version only ever tried .TW, so OTC names came back empty.
'=============================================================================
Function GetHistoricalData(Ticker As String, ByRef outDates() As Date, ByRef outPrices() As Double) As Boolean
    Dim t As String
    t = UCase$(Trim$(Ticker))

    If InStr(t, ".") = 0 And IsNumeric(t) Then
        If FetchYahooSeries(t & ".TW", outDates, outPrices) Then
            GetHistoricalData = True
            Exit Function
        End If
        GetHistoricalData = FetchYahooSeries(t & ".TWO", outDates, outPrices)
    Else
        GetHistoricalData = FetchYahooSeries(t, outDates, outPrices)
    End If
End Function

Private Function FetchYahooSeries(ByVal Ticker As String, ByRef outDates() As Date, _
                                  ByRef outPrices() As Double) As Boolean
    Dim http As Object
    Dim url As String, response As String
    Dim tsStr As String, closeStr As String
    Dim tsArr() As String, closeArr() As String
    Dim i As Long

    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?range=10y&interval=1d"

    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        response = .responseText
    End With
    On Error GoTo 0
    If Len(response) = 0 Then Exit Function

    Dim tsStart As Long, tsEnd As Long
    tsStart = InStr(response, """timestamp"":[") + 12
    tsEnd = InStr(tsStart, response, "]")
    If tsStart <= 12 Or tsEnd <= tsStart Then Exit Function
    tsStr = Mid(response, tsStart, tsEnd - tsStart)
    tsArr = Split(tsStr, ",")

    Dim closeStart As Long, closeEnd As Long
    closeStart = InStr(response, """adjclose"":[{""adjclose"":[") + 23
    If closeStart < 24 Then closeStart = InStr(response, """close"":[") + 8
    closeEnd = InStr(closeStart, response, "]")
    If closeStart <= 20 Or closeEnd <= closeStart Then Exit Function
    closeStr = Mid(response, closeStart, closeEnd - closeStart)
    closeArr = Split(closeStr, ",")

    Dim count As Long
    count = UBound(tsArr)
    If UBound(closeArr) < count Then count = UBound(closeArr)
    If count < 1 Then Exit Function

    ReDim outDates(0 To count)
    ReDim outPrices(0 To count)

    Dim validCount As Long
    validCount = 0
    For i = 0 To count
        If IsNumeric(tsArr(i)) And IsNumeric(closeArr(i)) Then
            outDates(validCount) = (CDbl(tsArr(i)) / 86400) + 25569
            outPrices(validCount) = CDbl(closeArr(i))
            validCount = validCount + 1
        End If
    Next i

    If validCount > 1 Then
        ReDim Preserve outDates(0 To validCount - 1)
        ReDim Preserve outPrices(0 To validCount - 1)
        FetchYahooSeries = True
    End If
End Function
