Attribute VB_Name = "RRG"
Option Explicit

' ================================================================
'  RELATIVE ROTATION GRAPH (nav code RG, recalc RG!) - sheet "RRG"
' ----------------------------------------------------------------
'  Prototype (2026-09-12): an Excel port of sector-rotation-system/
'  rrg_dynamic.py, snapshot only (no time slider).  The maths is kept
'  identical to the Python so the two can be checked against each other:
'
'    rs        = adjclose(ticker) / adjclose(SPY)           (daily)
'    rs_ratio  = 100 * rs / SMA(rs, 65)
'    rs_mom    = 100 * rs_ratio / SMA(rs_ratio, 20)         (unsmoothed ratio)
'    both      = EWM(span 3, adjust=True)  (pandas default)
'    tail      = last 66 rows, sampled every 5th trading day backwards
'                from the latest one -> 14 points (13 weeks)
'    quadrant  = ratio >= 100 ? (mom >= 100 ? LEADING : WEAKENING)
'                             : (mom >= 100 ? IMPROVING : LAGGING)
'
'  Universe = CorelationMatrix.SectorList() (config.py universe + SPY).
'  Prices come straight from the Yahoo chart API (adjclose, 1y) so the
'  workbook does not depend on the Python side at all.
'
'  Layout (page rows; the nav bar adds 3 on top - see NavOffset):
'    row 1   title + subtitle
'    row 2   hint
'    row 4   table header, rows 5..30 one ETF each (A..I)
'    K4      chart RRG_MAIN (scatter, one series per ETF)
'    row 33  quadrant membership, row 39.. notes
'    Z5..    tail dates (shared by every ETF - all are aligned to SPY days)
'    AA4..   tail data block (2 columns per ETF, 14 rows) - chart source
'
'  FOCUS: double-click a ticker in column A -> only that trail stays lit
'  (thicker line, every dot dated), every other ETF goes dark grey and
'  loses its label, the table row gets the orange background.  Double-
'  click the same ticker again, or the TICKER header, to show all.  The
'  sheet's Worksheet_BeforeDoubleClick (RR4/SheetRRG_Code.txt) is written
'  into the sheet module by BuildRRG itself the first time the sheet is
'  built (needs "Trust access to the VBA project object model").  The
'  focused ticker lives in the hidden sheet name RRGFOCUS.
' ================================================================

Private Const RS_WINDOW As Long = 65
Private Const MOM_WINDOW As Long = 20
Private Const EWM_SPAN As Long = 3
Private Const TAIL_WEEKS As Long = 13
Private Const TAIL_SPACING As Long = 5
Private Const TAIL_POINTS As Long = TAIL_WEEKS + 1
Private Const BENCH As String = "SPY"

Private Const TBL_HDR As Long = 4
Private Const TBL_FIRST As Long = 5
Private Const DATA_COL As Long = 27          ' AA
Private Const DATE_COL As Long = 26          ' Z
Private Const CHART_COL As Long = 11         ' K
Private Const FOCUS_MARK As String = "RRGFOCUS"
Private Const DIM_GREY As Long = 4210752     ' RGB(64,64,64)

Private Const NAN As Double = -1E+300

Private Function QuadColor(ByVal q As String) As Long
    Select Case q
        Case "LEADING":   QuadColor = RGB(46, 125, 50)
        Case "WEAKENING": QuadColor = RGB(249, 168, 37)
        Case "LAGGING":   QuadColor = RGB(198, 40, 40)
        Case Else:        QuadColor = RGB(21, 101, 192)     ' IMPROVING
    End Select
End Function

Private Function Quadrant(ByVal x As Double, ByVal y As Double) As String
    If x >= 100 Then
        Quadrant = IIf(y >= 100, "LEADING", "WEAKENING")
    Else
        Quadrant = IIf(y >= 100, "IMPROVING", "LAGGING")
    End If
End Function

' ----------------------------------------------------------------
' Yahoo chart API: timestamps + adjclose.  Returns the number of valid
' points; days where adjclose is null are dropped (both arrays stay aligned).
Private Function FetchAdjSeries(ByVal Ticker As String, ByRef days() As Long, ByRef px() As Double) As Long
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=1y"
    On Error Resume Next
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    On Error GoTo 0
    If http.Status <> 200 Then Exit Function
    Dim resp As String: resp = http.responseText

    Dim p1 As Long, p2 As Long
    p1 = InStr(resp, """timestamp"":[")
    If p1 = 0 Then Exit Function
    p1 = p1 + Len("""timestamp"":[")
    p2 = InStr(p1, resp, "]")
    Dim tsParts() As String: tsParts = Split(Mid(resp, p1, p2 - p1), ",")

    ' the numeric array is the SECOND "adjclose":[ (the first opens the object)
    p1 = InStr(resp, """adjclose"":[")
    If p1 = 0 Then Exit Function
    p1 = InStr(p1 + 1, resp, """adjclose"":[")
    If p1 = 0 Then Exit Function
    p1 = p1 + Len("""adjclose"":[")
    p2 = InStr(p1, resp, "]")
    Dim pxParts() As String: pxParts = Split(Mid(resp, p1, p2 - p1), ",")

    Dim n As Long: n = UBound(tsParts) + 1
    If UBound(pxParts) + 1 < n Then n = UBound(pxParts) + 1
    ReDim days(0 To n - 1): ReDim px(0 To n - 1)
    Dim k As Long, cnt As Long
    For k = 0 To n - 1
        If IsNumeric(Trim(pxParts(k))) And IsNumeric(Trim(tsParts(k))) Then
            days(cnt) = CLng(Val(tsParts(k)) \ 86400)
            px(cnt) = CDbl(Val(pxParts(k)))
            cnt = cnt + 1
        End If
    Next k
    FetchAdjSeries = cnt
End Function

' ----------------------------------------------------------------
Sub BuildRRG()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("RRG")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = "RRG"
    End If
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fail
    Call NavStrip(ws)

    ' ---- universe (SPY last in the list is the benchmark) ----
    Dim lst As Variant: lst = SectorList()
    Dim nAll As Long: nAll = UBound(lst) + 1
    Dim tickers() As String, labels() As String, groups() As String
    Dim i As Long, j As Long, k As Long, n As Long
    n = 0
    ReDim tickers(0 To nAll - 1): ReDim labels(0 To nAll - 1): ReDim groups(0 To nAll - 1)
    For i = 0 To nAll - 1
        If UCase(lst(i)(0)) <> BENCH Then
            tickers(n) = lst(i)(0): labels(n) = lst(i)(1): groups(n) = lst(i)(2)
            n = n + 1
        End If
    Next i
    ReDim Preserve tickers(0 To n - 1): ReDim Preserve labels(0 To n - 1): ReDim Preserve groups(0 To n - 1)

    ' ---- benchmark ----
    Dim bDays() As Long, bPx() As Double, nb As Long
    Application.StatusBar = "RRG: fetching " & BENCH
    nb = FetchAdjSeries(BENCH, bDays, bPx)
    If nb < RS_WINDOW + MOM_WINDOW + 5 Then Err.Raise vbObjectError + 1, , "not enough " & BENCH & " data (" & nb & " days)"
    Dim bIdx As Object: Set bIdx = CreateObject("Scripting.Dictionary")
    For i = 0 To nb - 1: bIdx(bDays(i)) = i: Next i

    ' ---- per ticker: aligned rs -> ratio / momentum -> tail ----
    Dim tailX() As Double, tailY() As Double, tailN() As Long
    ReDim tailX(0 To n - 1, 0 To TAIL_POINTS - 1)
    ReDim tailY(0 To n - 1, 0 To TAIL_POINTS - 1)
    ReDim tailN(0 To n - 1)
    Dim rs() As Double, ratio() As Double, mom() As Double, sr() As Double, sm() As Double
    Dim vx() As Double, vy() As Double, nv As Long
    Dim tDays() As Long, tPx() As Double, nt As Long
    For i = 0 To n - 1
        Application.StatusBar = "RRG: fetching [" & (i + 1) & "/" & n & "] " & tickers(i)
        DoEvents
        nt = FetchAdjSeries(tickers(i), tDays, tPx)
        ReDim rs(0 To nb - 1): ReDim ratio(0 To nb - 1): ReDim mom(0 To nb - 1)
        ReDim sr(0 To nb - 1): ReDim sm(0 To nb - 1)
        For k = 0 To nb - 1: rs(k) = NAN: ratio(k) = NAN: mom(k) = NAN: sr(k) = NAN: sm(k) = NAN: Next k
        For k = 0 To nt - 1
            If bIdx.Exists(tDays(k)) Then rs(bIdx(tDays(k))) = tPx(k) / bPx(bIdx(tDays(k)))
        Next k
        Call RollingRatio(rs, ratio, RS_WINDOW)
        Call RollingRatio(ratio, mom, MOM_WINDOW)
        Call EwmAdjust(ratio, sr, EWM_SPAN)
        Call EwmAdjust(mom, sm, EWM_SPAN)
        ' rows where both are valid (pandas dropna), keep the last 66
        ReDim vx(0 To nb - 1): ReDim vy(0 To nb - 1): nv = 0
        For k = 0 To nb - 1
            If sr(k) <> NAN And sm(k) <> NAN Then vx(nv) = sr(k): vy(nv) = sm(k): nv = nv + 1
        Next k
        Dim span As Long: span = TAIL_WEEKS * TAIL_SPACING + 1
        If nv < span Then span = nv
        ' sample backwards from the newest point every TAIL_SPACING rows
        Dim cnt As Long: cnt = 0
        For k = nv - 1 To nv - span Step -TAIL_SPACING
            cnt = cnt + 1
        Next k
        tailN(i) = cnt
        j = TAIL_POINTS - 1
        For k = nv - 1 To nv - span Step -TAIL_SPACING
            tailX(i, j) = vx(k): tailY(i, j) = vy(k): j = j - 1
        Next k
    Next i
    Application.StatusBar = False

    ' ---- page ----
    Application.ScreenUpdating = False
    Call ShapesMoveOnly(ws)
    Dim co As ChartObject
    For Each co In ws.ChartObjects: co.Delete: Next co
    ws.cells.Clear
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.DisplayGridlines = False
    Dim rr As Long
    For rr = 1 To 60: ws.Rows(rr).RowHeight = 18: Next rr
    ws.Columns(1).ColumnWidth = 8: ws.Columns(2).ColumnWidth = 9: ws.Columns(3).ColumnWidth = 15
    ws.Columns(4).ColumnWidth = 9: ws.Columns(5).ColumnWidth = 9: ws.Columns(6).ColumnWidth = 11
    ws.Columns(7).ColumnWidth = 9: ws.Columns(8).ColumnWidth = 9: ws.Columns(9).ColumnWidth = 7
    ws.Columns(10).ColumnWidth = 3

    Dim asOf As Date: asOf = DateSerial(1970, 1, 1) + bDays(nb - 1)
    With ws.cells(1, 1)
        .Value = "RELATIVE ROTATION GRAPH"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 14
    End With
    ws.Rows(1).RowHeight = 24
    With ws.cells(1, 4)
        .Value = n & " ETFs vs " & BENCH & "  .  daily adjclose  .  RS " & RS_WINDOW & "d / MOM " & MOM_WINDOW & _
                 "d / EWM " & EWM_SPAN & "  .  " & TAIL_WEEKS & "-week tail  .  as of " & Format(asOf, "yyyy/mm/dd") & _
                 "  .  built " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RGB(120, 120, 120): .Font.Size = 9
    End With
    With ws.cells(2, 1)
        .Value = "RG! <GO> refetches and redraws  .  double-click a ticker to focus it (again, or the TICKER header, to clear)  .  same maths as rrg_dynamic.py"
        .Font.Color = RGB(120, 120, 120): .Font.Size = 9
    End With
    With ws.Range(ws.cells(2, 1), ws.cells(2, 9)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With

    ' ---- table ----
    Dim hdr As Variant
    hdr = Array("TICKER", "LABEL", "GROUP", "RS-RATIO", "RS-MOM", "QUADRANT", "1W dRAT", "1W dMOM", "PTS")
    For j = 0 To UBound(hdr)
        With ws.cells(TBL_HDR, j + 1)
            .Value = hdr(j)
            .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 9
            .HorizontalAlignment = IIf(j >= 3, xlRight, xlLeft)
        End With
    Next j
    With ws.Range(ws.cells(TBL_HDR, 1), ws.cells(TBL_HDR, 9)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With
    Dim quadList(0 To 3) As String
    quadList(0) = "LEADING": quadList(1) = "IMPROVING": quadList(2) = "WEAKENING": quadList(3) = "LAGGING"
    Dim quadMembers(0 To 3) As String
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    xMin = 99: xMax = 101: yMin = 99: yMax = 101
    For i = 0 To n - 1
        Dim r As Long: r = TBL_FIRST + i
        ws.cells(r, 1).Value = tickers(i)
        ws.cells(r, 2).Value = labels(i)
        ws.cells(r, 3).Value = groups(i)
        If tailN(i) = 0 Then
            ws.cells(r, 6).Value = "NO DATA"
        Else
            Dim lx As Double, ly As Double
            lx = tailX(i, TAIL_POINTS - 1): ly = tailY(i, TAIL_POINTS - 1)
            Dim q As String: q = Quadrant(lx, ly)
            ws.cells(r, 4).Value = lx: ws.cells(r, 5).Value = ly
            ws.Range(ws.cells(r, 4), ws.cells(r, 5)).NumberFormat = "0.00"
            ws.cells(r, 6).Value = q
            If tailN(i) >= 2 Then
                ws.cells(r, 7).Value = lx - tailX(i, TAIL_POINTS - 2)
                ws.cells(r, 8).Value = ly - tailY(i, TAIL_POINTS - 2)
                ws.Range(ws.cells(r, 7), ws.cells(r, 8)).NumberFormat = "+0.00;-0.00;0.00"
            End If
            ws.cells(r, 9).Value = tailN(i)
            For j = 0 To 3
                If quadList(j) = q Then quadMembers(j) = quadMembers(j) & IIf(quadMembers(j) = "", "", "  ") & tickers(i)
            Next j
            ' axis range over every tail point (python: min/max of all in-range rows + 1.5 pad)
            For k = TAIL_POINTS - tailN(i) To TAIL_POINTS - 1
                If tailX(i, k) < xMin Then xMin = tailX(i, k)
                If tailX(i, k) > xMax Then xMax = tailX(i, k)
                If tailY(i, k) < yMin Then yMin = tailY(i, k)
                If tailY(i, k) > yMax Then yMax = tailY(i, k)
            Next k
        End If
        ws.Range(ws.cells(r, 4), ws.cells(r, 9)).HorizontalAlignment = xlRight
        ws.cells(r, 6).HorizontalAlignment = xlRight
    Next i
    Dim lastRow As Long: lastRow = TBL_FIRST + n - 1
    With ws.Range(ws.cells(lastRow, 1), ws.cells(lastRow, 9)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With
    Call PaintTableRows(ws, "")

    ' ---- quadrant membership ----
    Dim qr As Long: qr = lastRow + 2
    With ws.cells(qr, 1)
        .Value = "QUADRANTS": .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With
    For j = 0 To 3
        With ws.cells(qr + 1 + j, 1)
            .Value = quadList(j): .Font.Color = QuadColor(quadList(j)): .Font.Bold = True
        End With
        ws.cells(qr + 1 + j, 3).Value = IIf(quadMembers(j) = "", "-", quadMembers(j))
    Next j

    ' ---- notes ----
    Dim nr As Long: nr = qr + 6
    With ws.cells(nr, 1)
        .Value = "HOW THESE ARE COMPUTED": .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With
    Dim notes As Variant
    notes = Array( _
        "RS-RATIO  = 100 x (px / SPY) / SMA65(px / SPY), then EWM span 3  -  relative strength vs its own quarter trend", _
        "RS-MOM    = 100 x RS-RATIO / SMA20(RS-RATIO), then EWM span 3  -  is that strength accelerating (>100) or fading", _
        "TAIL      = the last 13 weeks, one dot per 5 trading days; the big dot is today, older dots fade", _
        "QUADRANT  = LEADING (both >= 100) / WEAKENING (ratio >= 100, mom < 100) / LAGGING (both < 100) / IMPROVING (ratio < 100, mom >= 100)", _
        "1W dRAT / dMOM = change since the previous tail dot (5 trading days), the direction the tail is heading", _
        "Same windows and smoothing as sector-rotation-system/rrg_dynamic.py; prices are Yahoo adjclose, so dividends are in.")
    For j = 0 To UBound(notes)
        ws.cells(nr + 1 + j, 1).Value = notes(j)
        ws.cells(nr + 1 + j, 1).Font.Color = RGB(150, 150, 150)
        ws.cells(nr + 1 + j, 1).Font.Size = 8
    Next j

    ' ---- tail data block (chart source) ----
    With ws.cells(TBL_HDR - 1, DATA_COL)
        .Value = "TAIL DATA (chart source, oldest -> newest)": .Font.Color = RGB(90, 90, 90): .Font.Size = 8
    End With
    ' tail dates: every ETF is aligned to SPY's days and sampled backwards
    ' from the same last day, so one date column serves all of them
    ws.cells(TBL_HDR, DATE_COL).Value = "date": ws.cells(TBL_HDR, DATE_COL).Font.Color = RGB(90, 90, 90)
    For k = 0 To TAIL_POINTS - 1
        Dim di As Long: di = nb - 1 - TAIL_SPACING * (TAIL_POINTS - 1 - k)
        If di >= 0 Then ws.cells(TBL_FIRST + k, DATE_COL).Value = DateSerial(1970, 1, 1) + bDays(di)
    Next k
    ws.Range(ws.cells(TBL_FIRST, DATE_COL), ws.cells(TBL_FIRST + TAIL_POINTS - 1, DATE_COL)).NumberFormat = "mm/dd"
    ws.Range(ws.cells(TBL_FIRST, DATE_COL), ws.cells(TBL_FIRST + TAIL_POINTS - 1, DATE_COL)).Font.Color = RGB(90, 90, 90)
    ws.Columns(DATE_COL).ColumnWidth = 7
    For i = 0 To n - 1
        Dim cx As Long: cx = DATA_COL + 2 * i
        ws.cells(TBL_HDR, cx).Value = tickers(i): ws.cells(TBL_HDR, cx + 1).Value = "mom"
        ws.Range(ws.cells(TBL_HDR, cx), ws.cells(TBL_HDR, cx + 1)).Font.Color = RGB(90, 90, 90)
        For k = 0 To TAIL_POINTS - 1
            If k >= TAIL_POINTS - tailN(i) Then
                ws.cells(TBL_FIRST + k, cx).Value = tailX(i, k)
                ws.cells(TBL_FIRST + k, cx + 1).Value = tailY(i, k)
            End If
        Next k
        ws.Range(ws.cells(TBL_FIRST, cx), ws.cells(TBL_FIRST + TAIL_POINTS - 1, cx + 1)).NumberFormat = "0.00"
        ws.Range(ws.cells(TBL_FIRST, cx), ws.cells(TBL_FIRST + TAIL_POINTS - 1, cx + 1)).Font.Color = RGB(90, 90, 90)
        ws.Columns(cx).ColumnWidth = 7: ws.Columns(cx + 1).ColumnWidth = 7
    Next i

    ' ---- chart ----
    Call DrawRrgChart(ws, tickers, tailX, tailY, tailN, n, xMin - 1.5, xMax + 1.5, yMin - 1.5, yMax + 1.5, asOf)

    Call NavAdd(ws, "RG")
    Call SetFocusMark(ws, "")
    Call EnsureSheetCode(ws)
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    Call NavNotify("RRG rebuilt: " & n & " ETFs vs " & BENCH & " as of " & Format(asOf, "yyyy/mm/dd"))
    Exit Sub
Fail:
    Dim failMsg As String: failMsg = Err.Description
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    On Error Resume Next
    Call NavAdd(ws, "RG")
    Call NavNotify("RRG failed: " & failMsg, True)
End Sub

' 100 * x / SMA(x, w); NaN unless the whole window is valid (pandas min_periods = window)
Private Sub RollingRatio(ByRef src() As Double, ByRef dst() As Double, ByVal w As Long)
    Dim i As Long, k As Long, s As Double, ok As Boolean
    For i = LBound(src) To UBound(src)
        dst(i) = NAN
        If i - LBound(src) + 1 >= w Then
            s = 0: ok = True
            For k = i - w + 1 To i
                If src(k) = NAN Then ok = False: Exit For
                s = s + src(k)
            Next k
            If ok And s <> 0 Then dst(i) = 100 * src(i) / (s / w)
        End If
    Next i
End Sub

' pandas Series.ewm(span=n).mean() with adjust=True (the default): weighted
' average with weights (1-a)^k, a = 2/(n+1).  Leading NaNs stay NaN.
Private Sub EwmAdjust(ByRef src() As Double, ByRef dst() As Double, ByVal spanN As Long)
    Dim a As Double: a = 2 / (spanN + 1)
    Dim num As Double, den As Double, started As Boolean
    Dim i As Long
    For i = LBound(src) To UBound(src)
        If src(i) = NAN Then
            dst(i) = NAN
            If started Then num = num * (1 - a): den = den * (1 - a)   ' ignore_na=False: the gap still decays the weights
        Else
            num = num * (1 - a) + src(i)
            den = den * (1 - a) + 1
            started = True
            dst(i) = num / den
        End If
    Next i
End Sub

Private Sub DrawRrgChart(ws As Worksheet, tickers() As String, tailX() As Double, tailY() As Double, tailN() As Long, _
                         ByVal n As Long, ByVal x0 As Double, ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double, ByVal asOf As Date)
    Dim topRow As Long: topRow = TBL_HDR + NavOffset(ws)
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(ws.Columns(CHART_COL).Left, ws.Rows(topRow).Top, 560, 470)
    co.Name = "RRG_MAIN"
    co.Placement = xlMove
    Dim ch As Chart: Set ch = co.Chart
    ch.ChartType = xlXYScatterLines
    Do While ch.SeriesCollection.count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.HasLegend = False
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(8, 8, 8)
    ch.PlotArea.Format.Line.Visible = msoFalse
    ch.HasTitle = True
    ch.ChartTitle.Text = "RRG  vs " & BENCH & "   as of " & Format(asOf, "yyyy/mm/dd")
    With ch.ChartTitle.Format.TextFrame2.TextRange.Font
        .Name = "Consolas": .Size = 10: .Bold = msoTrue: .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    Dim i As Long, k As Long
    For i = 0 To n - 1
        If tailN(i) > 0 Then
            Dim cx As Long: cx = DATA_COL + 2 * i
            Dim firstRow As Long: firstRow = TBL_FIRST + NavOffset(ws) + (TAIL_POINTS - tailN(i))
            Dim lastRow As Long: lastRow = TBL_FIRST + NavOffset(ws) + TAIL_POINTS - 1
            Dim s As Series: Set s = ch.SeriesCollection.NewSeries
            s.Name = tickers(i)
            s.XValues = ws.Range(ws.cells(firstRow, cx), ws.cells(lastRow, cx))
            s.Values = ws.Range(ws.cells(firstRow, cx + 1), ws.cells(lastRow, cx + 1))
            Call StyleSeries(ws, s, QuadColor(Quadrant(tailX(i, TAIL_POINTS - 1), tailY(i, TAIL_POINTS - 1))), tailN(i), 0)
        End If
    Next i

    Dim ax As Axis
    Set ax = ch.Axes(xlCategory)
    ax.MinimumScale = x0: ax.MaximumScale = x1
    ax.CrossesAt = 100
    ax.HasMajorGridlines = False
    ax.TickLabelPosition = xlTickLabelPositionLow
    ax.Format.Line.ForeColor.RGB = RGB(110, 110, 110)
    ax.TickLabels.Font.Name = "Consolas": ax.TickLabels.Font.Size = 8: ax.TickLabels.Font.Color = RGB(150, 150, 150)
    ax.TickLabels.NumberFormat = "0"
    ax.HasTitle = True: ax.AxisTitle.Text = "RS-RATIO"
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Consolas"
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 8
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(180, 180, 180)
    Set ax = ch.Axes(xlValue)
    ax.MinimumScale = y0: ax.MaximumScale = y1
    ax.CrossesAt = 100
    ax.HasMajorGridlines = False
    ax.TickLabelPosition = xlTickLabelPositionLow
    ax.Format.Line.ForeColor.RGB = RGB(110, 110, 110)
    ax.TickLabels.Font.Name = "Consolas": ax.TickLabels.Font.Size = 8: ax.TickLabels.Font.Color = RGB(150, 150, 150)
    ax.TickLabels.NumberFormat = "0"
    ax.HasTitle = True: ax.AxisTitle.Text = "RS-MOMENTUM"
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Consolas"
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 8
    ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(180, 180, 180)

    ' quadrant labels in the four corners of the plot area
    Dim pl As Double, pt As Double, pw As Double, ph As Double
    pl = ch.PlotArea.InsideLeft: pt = ch.PlotArea.InsideTop
    pw = ch.PlotArea.InsideWidth: ph = ch.PlotArea.InsideHeight
    Call QuadLabel(ch, "LEADING", pl + pw - 78, pt + 4, msoAnchorTop)
    Call QuadLabel(ch, "WEAKENING", pl + pw - 78, pt + ph - 18, msoAnchorBottom)
    Call QuadLabel(ch, "LAGGING", pl + 4, pt + ph - 18, msoAnchorBottom)
    Call QuadLabel(ch, "IMPROVING", pl + 4, pt + 4, msoAnchorTop)
End Sub

' ----------------------------------------------------------------
' Series look.  mode 0 = normal (quadrant colour, ticker label on the last
' dot), 1 = dimmed (dark grey, no label), 2 = focused (thicker, every dot
' dated from the Z column).  Per-point fill transparency would reset the
' marker colour to the theme's, so fading is done by blending towards black.
Private Sub StyleSeries(ws As Worksheet, s As Series, ByVal qc As Long, ByVal np As Long, ByVal mode As Long)
    Dim k As Long
    Dim baseCol As Long: baseCol = IIf(mode = 1, DIM_GREY, qc)
    s.Format.Line.ForeColor.RGB = baseCol
    s.Format.Line.Weight = IIf(mode = 2, 2, 1)
    s.Format.Line.Transparency = IIf(mode = 2, 0.15, 0.5)
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 4
    s.MarkerBackgroundColor = baseCol
    s.MarkerForegroundColor = baseCol
    s.HasDataLabels = False
    For k = 1 To np - 1
        Dim fade As Double: fade = 0.25 + 0.75 * (k - 1) / (np - 1)
        s.Points(k).MarkerSize = IIf(mode = 2, 5, 3)
        s.Points(k).MarkerBackgroundColor = IIf(mode = 1, DIM_GREY, Dim2(qc, fade))
        s.Points(k).MarkerForegroundColor = IIf(mode = 1, DIM_GREY, Dim2(qc, fade))
    Next k
    With s.Points(np)
        .MarkerSize = IIf(mode = 1, 6, IIf(mode = 2, 11, 9))
        .MarkerBackgroundColor = baseCol
        .MarkerForegroundColor = baseCol
    End With
    If mode = 0 Then
        Call PointLabel(s.Points(np), s.Name, RGB(210, 210, 210), 8)
    ElseIf mode = 2 Then
        Dim off As Long: off = NavOffset(ws)
        For k = 1 To np
            Dim dr As Long: dr = TBL_FIRST + off + TAIL_POINTS - np + k - 1
            Dim txt As String: txt = Format(ws.cells(dr, DATE_COL).Value, "mm/dd")
            If k = np Then txt = s.Name & " " & txt
            Call PointLabel(s.Points(k), txt, IIf(k = np, RGB(255, 255, 255), RGB(190, 190, 190)), IIf(k = np, 9, 7))
        Next k
    End If
End Sub

Private Sub PointLabel(p As Point, ByVal txt As String, ByVal col As Long, ByVal sz As Long)
    p.HasDataLabel = True
    With p.DataLabel
        ' turning every Show* flag off deletes the label (and .Text then
        ' fails with "invalid parameter"), so keep the series name on and
        ' overwrite the text
        .ShowValue = False
        .ShowCategoryName = False
        .ShowSeriesName = True
        .Text = txt
        .Position = xlLabelPositionRight
        .Format.TextFrame2.TextRange.Font.Name = "Consolas"
        .Format.TextFrame2.TextRange.Font.Size = sz
        .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = col
    End With
End Sub

' Table colours for every ETF row.  focusTk = "" -> normal look; otherwise
' that row gets the orange background and every other row goes dim.
Private Sub PaintTableRows(ws As Worksheet, ByVal focusTk As String)
    Dim off As Long: off = NavOffset(ws)
    Dim r As Long, i As Long
    i = 0
    r = TBL_FIRST + off
    Do While ws.cells(r, 1).Value <> ""
        Dim rng As Range: Set rng = ws.Range(ws.cells(r, 1), ws.cells(r, 9))
        Dim tk As String: tk = CStr(ws.cells(r, 1).Value)
        Dim q As String: q = CStr(ws.cells(r, 6).Value)
        rng.Font.Bold = False
        If focusTk <> "" And tk = focusTk Then
            rng.Interior.Color = RR4_ACCENT
            rng.Font.Color = RGB(0, 0, 0)
            rng.Font.Bold = True
        ElseIf focusTk <> "" Then
            rng.Interior.Color = IIf(i Mod 2 = 0, RGB(8, 8, 8), RGB(14, 14, 14))
            rng.Font.Color = RGB(75, 75, 75)
        Else
            rng.Interior.Color = IIf(i Mod 2 = 0, RGB(8, 8, 8), RGB(14, 14, 14))
            rng.Font.Color = RGB(221, 221, 221)
            ws.cells(r, 1).Font.Color = RR4_ACCENT: ws.cells(r, 1).Font.Bold = True
            ws.cells(r, 3).Font.Color = RGB(150, 150, 150)
            If q = "NO DATA" Then
                ws.cells(r, 6).Font.Color = RGB(120, 120, 120)
            Else
                ws.cells(r, 6).Font.Color = QuadColor(q): ws.cells(r, 6).Font.Bold = True
            End If
            If ws.cells(r, 7).Value <> "" Then
                ws.cells(r, 7).Font.Color = IIf(ws.cells(r, 7).Value >= 0, RGB(220, 80, 80), RGB(80, 200, 120))
                ws.cells(r, 8).Font.Color = IIf(ws.cells(r, 8).Value >= 0, RGB(220, 80, 80), RGB(80, 200, 120))
            End If
            ws.cells(r, 9).Font.Color = RGB(120, 120, 120)
        End If
        i = i + 1: r = r + 1
    Loop
End Sub

' ----------------------------------------------------------------
' Focus: "" shows every ETF, otherwise only tk stays lit.
Public Sub RrgFocus(ws As Worksheet, ByVal tk As String)
    Dim co As ChartObject
    On Error Resume Next
    Set co = ws.ChartObjects("RRG_MAIN")
    On Error GoTo 0
    If co Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    ' quadrant per ticker from the table (the chart series carry no colour memory)
    Dim off As Long: off = NavOffset(ws)
    Dim quadOf As Object: Set quadOf = CreateObject("Scripting.Dictionary")
    Dim r As Long: r = TBL_FIRST + off
    Do While ws.cells(r, 1).Value <> ""
        quadOf(CStr(ws.cells(r, 1).Value)) = CStr(ws.cells(r, 6).Value)
        r = r + 1
    Loop
    Dim s As Series
    For Each s In co.Chart.SeriesCollection
        Dim mode As Long
        If tk = "" Then mode = 0 Else mode = IIf(s.Name = tk, 2, 1)
        Call StyleSeries(ws, s, QuadColor(CStr(quadOf(s.Name))), s.Points.count, mode)
    Next s
    Call PaintTableRows(ws, tk)
    Call SetFocusMark(ws, tk)
    Application.ScreenUpdating = True
    If tk = "" Then
        Call NavNotify("RRG: showing all")
    Else
        Call NavNotify("RRG: focus " & tk & " (" & quadOf(tk) & ")  -  double-click it again or the TICKER header to show all")
    End If
End Sub

' Same, by sheet name - callable from the Immediate window / Application.Run.
Public Sub RrgFocusByName(ByVal tk As String)
    Call RrgFocus(ThisWorkbook.Sheets("RRG"), UCase(Trim(tk)))
End Sub

' Sheet double-click handler (called from SheetRRG_Code.txt).
Public Sub RrgDoubleClick(ws As Worksheet, ByVal Target As Range, ByRef Cancel As Boolean)
    If Target.Column <> 1 Then Exit Sub
    Dim off As Long: off = NavOffset(ws)
    Dim tk As String: tk = Trim(CStr(Target.Value))
    If Target.Row = TBL_HDR + off Then
        Cancel = True
        Call RrgFocus(ws, "")
    ElseIf Target.Row > TBL_HDR + off And tk <> "" Then
        If ws.cells(Target.Row, 6).Value = "" Then Exit Sub        ' below the table
        Cancel = True
        If GetFocusMark(ws) = tk Then Call RrgFocus(ws, "") Else Call RrgFocus(ws, tk)
    End If
End Sub

Private Sub SetFocusMark(ws As Worksheet, ByVal tk As String)
    On Error Resume Next
    ws.Names(FOCUS_MARK).Delete
    On Error GoTo 0
    ws.Names.Add Name:=FOCUS_MARK, RefersTo:="=""" & tk & """", Visible:=False
End Sub

Private Function GetFocusMark(ws As Worksheet) As String
    On Error Resume Next
    Dim v As String: v = ws.Names(FOCUS_MARK).RefersTo      ' ="SMH"
    On Error GoTo 0
    v = Replace(Replace(v, "=", ""), """", "")
    GetFocusMark = v
End Function

' Write the sheet's event code (a copy of RR4/SheetRRG_Code.txt) into the
' document module if it is not there yet.  The RRG sheet is created by
' BuildRRG, so the code cannot be injected ahead of time.
Private Sub EnsureSheetCode(ws As Worksheet)
    On Error GoTo Skip
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 Then
            If comp.Properties("Name").Value = ws.Name Then
                Dim cm As Object: Set cm = comp.CodeModule
                If cm.CountOfLines > 0 Then
                    If InStr(cm.Lines(1, cm.CountOfLines), "RrgDoubleClick") > 0 Then Exit Sub
                    cm.DeleteLines 1, cm.CountOfLines
                End If
                cm.AddFromString "Option Explicit" & vbCrLf & vbCrLf & _
                    "' RRG page: double-click a ticker in column A to focus it (see RRG.bas)" & vbCrLf & _
                    "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & vbCrLf & _
                    "    Call RrgDoubleClick(Me, Target, Cancel)" & vbCrLf & _
                    "End Sub" & vbCrLf
                Exit Sub
            End If
        End If
    Next comp
    Exit Sub
Skip:
    Call NavNotify("RRG built, but the double-click code could not be written (enable Trust access to the VBA project object model, or paste RR4/SheetRRG_Code.txt)", True)
End Sub

' colour scaled towards black: f = 1 keeps it, f = 0.25 is a quarter as bright
Private Function Dim2(ByVal c As Long, ByVal f As Double) As Long
    Dim2 = RGB(Int((c Mod 256) * f), Int(((c \ 256) Mod 256) * f), Int(((c \ 65536) Mod 256) * f))
End Function

Private Sub QuadLabel(ch As Chart, ByVal txt As String, ByVal l As Double, ByVal t As Double, ByVal anchor As Long)
    Dim shp As Shape
    Set shp = ch.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, 74, 14)
    shp.Name = "RRG_Q_" & txt
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Name = "Consolas"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = QuadColor(txt)
        .TextRange.ParagraphFormat.Alignment = IIf(txt = "LAGGING" Or txt = "IMPROVING", msoAlignLeft, msoAlignRight)
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub
