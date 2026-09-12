Attribute VB_Name = "CorelationMatrix"
Option Explicit

' ================================================================
'  SECTOR CORRELATION (nav code CC) - sheet "Correlation"
' ----------------------------------------------------------------
'  v2 (2026-09-12): RR4 page look (orange accent, grey dividers, the
'  continuous red/green heat map shared with HoldingsCorr), and the ETF
'  universe now mirrors sector-rotation-system/config.py (RotationConfig
'  .universe + SPY benchmark) grouped by its TICKER_CATEGORY - keep the
'  two lists in sync by hand when that config changes.
'
'  Layout (page rows; the nav bar adds 3 on top - see NavOffset):
'    row 1   title + subtitle
'    row 2   A "DAYS <GO>"  B input (the sheet code re-runs on B2 edits)
'    row 4   header: column labels (short names)
'    row 5.. one row per ETF: A label, B ticker, matrix from column D,
'            AVG CORR and CORR TO SPY after the matrix
'    below   colour scale, group averages, top / bottom pairs, notes
' ================================================================

Private Const MATRIX_ROW As Long = 4
Private Const MATRIX_COL As Long = 4         ' D
Private Const DEFAULT_DAYS As Long = 60

' Universe: ticker | short label | group. Order = group order.
Private Function SectorList() As Variant
    SectorList = Array( _
        Array("SMH", "SMH", "SEMI"), Array("SOXX", "SOXX", "SEMI"), _
        Array("CHAT", "GenAI", "AI/TECH"), Array("MAGS", "Mag7", "AI/TECH"), Array("SKYY", "Cloud", "AI/TECH"), _
        Array("IGV", "Softwr", "AI/TECH"), Array("CIBR", "Cyber", "AI/TECH"), Array("DRAM", "DRAM", "AI/TECH"), _
        Array("QQQ", "QQQ", "INDEX/FACTOR"), Array("VTV", "Value", "INDEX/FACTOR"), Array("SPMO", "Momntm", "INDEX/FACTOR"), _
        Array("XLK", "Tech", "SECTOR SPDR"), Array("XLF", "Finance", "SECTOR SPDR"), Array("XLE", "Energy", "SECTOR SPDR"), _
        Array("XLV", "Health", "SECTOR SPDR"), Array("XLI", "Indust", "SECTOR SPDR"), Array("XLY", "ConDisc", "SECTOR SPDR"), _
        Array("XLP", "ConStap", "SECTOR SPDR"), Array("XLU", "Utility", "SECTOR SPDR"), Array("XLB", "Materl", "SECTOR SPDR"), _
        Array("XLRE", "RealEst", "SECTOR SPDR"), Array("XLC", "CommSvc", "SECTOR SPDR"), _
        Array("CPER", "Copper", "COMMODITY"), Array("GLD", "Gold", "COMMODITY"), _
        Array("ITA", "Defense", "SINGLE INDUSTRY"), Array("CRAK", "Refiner", "SINGLE INDUSTRY"), _
        Array("SPY", "SPY", "BENCHMARK"))
End Function

Sub BuildCorrelationMatrix()
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("Correlation")
    On Error GoTo 0
    If wsC Is Nothing Then
        Set wsC = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsC.Name = "Correlation"
    End If
    ' Writing the DAYS cell below would fire the sheet's Worksheet_Change,
    ' which calls this routine again re-entrantly (the inner run inserted the
    ' bar rows while the outer one kept writing unshifted - the page came out
    ' scrambled). Events off for the whole build.
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fail
    Call NavStrip(wsC)      ' nav bar rows off: the page is drawn from row 1 and B2 is read below

    ' days: read BEFORE the clear (typed input)
    Dim nDays As Long: nDays = DEFAULT_DAYS
    If IsNumeric(wsC.cells(2, 2).Value) And wsC.cells(2, 2).Value > 0 Then nDays = CLng(wsC.cells(2, 2).Value)
    If nDays < 10 Then nDays = 10
    If nDays > 240 Then nDays = 240      ' FetchPriceArray pulls 1y

    Dim lst As Variant: lst = SectorList()
    Dim nSec As Long: nSec = UBound(lst) + 1
    Dim tickers() As String, labels() As String, groups() As String
    ReDim tickers(0 To nSec - 1): ReDim labels(0 To nSec - 1): ReDim groups(0 To nSec - 1)
    Dim i As Long, j As Long
    For i = 0 To nSec - 1
        tickers(i) = lst(i)(0): labels(i) = lst(i)(1): groups(i) = lst(i)(2)
    Next i

    ' ---- base style ----
    Application.ScreenUpdating = False
    wsC.cells.Clear
    With wsC.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 8
        .VerticalAlignment = xlCenter
    End With
    wsC.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.DisplayGridlines = False
    Dim rr As Long
    For rr = 1 To 80: wsC.Rows(rr).RowHeight = 17: Next rr

    ' ---- title + input ----
    With wsC.cells(1, 1)
        .Value = "SECTOR CORRELATION"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 14
    End With
    wsC.Rows(1).RowHeight = 24
    With wsC.cells(1, MATRIX_COL)
        .Value = nSec & " ETFs (sector-rotation universe + SPY)  .  " & nDays & "-day daily returns  .  updated " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
    End With
    With wsC.cells(2, 1)
        .Value = "DAYS <GO>"
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 10
    End With
    With wsC.cells(2, 2)
        .Value = nDays
        .NumberFormat = "0"
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
    End With
    With wsC.cells(2, MATRIX_COL)
        .Value = "type a window (10-240 trading days) + Enter to recompute"
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
    End With
    Dim lastCol As Long: lastCol = MATRIX_COL + nSec + 2      ' gap, AVG, SPY
    With wsC.Range(wsC.cells(2, 1), wsC.cells(2, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    ' ---- fetch returns ----
    Dim retData() As Double
    ReDim retData(0 To nSec - 1, 0 To nDays - 2)
    Dim ok() As Boolean: ReDim ok(0 To nSec - 1)
    For i = 0 To nSec - 1
        Application.StatusBar = "Fetching [" & (i + 1) & "/" & nSec & "] " & tickers(i)
        DoEvents
        Dim prices() As Double
        prices = FetchPriceArray(tickers(i), nDays)
        Dim nz As Long: nz = 0
        For j = 0 To nDays - 2
            If j + 1 <= UBound(prices) And prices(j) > 0 And prices(j + 1) > 0 Then
                retData(i, j) = (prices(j + 1) - prices(j)) / prices(j)
                nz = nz + 1
            Else
                retData(i, j) = 0
            End If
        Next j
        ok(i) = (nz >= 5)
    Next i

    ' ---- correlation matrix ----
    Dim corr() As Double: ReDim corr(0 To nSec - 1, 0 To nSec - 1)
    For i = 0 To nSec - 1
        For j = 0 To nSec - 1
            If i = j Then
                corr(i, j) = 1
            ElseIf ok(i) And ok(j) Then
                corr(i, j) = CalcCorrelation(retData, i, j, nDays - 1)
            End If
        Next j
    Next i
    Dim spyIdx As Long: spyIdx = -1
    For i = 0 To nSec - 1
        If tickers(i) = "SPY" Then spyIdx = i
    Next i

    ' ---- columns ----
    wsC.Columns(1).ColumnWidth = 12
    wsC.Columns(2).ColumnWidth = 7
    wsC.Columns(3).ColumnWidth = 1
    For i = 0 To nSec - 1
        wsC.Columns(MATRIX_COL + i).ColumnWidth = 7.5     ' fits a 7-char label horizontally
    Next i
    wsC.Columns(MATRIX_COL + nSec).ColumnWidth = 1
    wsC.Columns(MATRIX_COL + nSec + 1).ColumnWidth = 8
    wsC.Columns(MATRIX_COL + nSec + 2).ColumnWidth = 8

    ' ---- header row ----
    Dim hdrCells As Variant: hdrCells = Array("SECTOR", "TICKER")
    For i = 0 To 1
        With wsC.cells(MATRIX_ROW, i + 1)
            .Value = hdrCells(i)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = IIf(i = 0, xlLeft, xlCenter)
        End With
    Next i
    For i = 0 To nSec - 1
        With wsC.cells(MATRIX_ROW, MATRIX_COL + i)
            .Value = labels(i)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
            .Font.Size = 7
        End With
    Next i
    wsC.Rows(MATRIX_ROW).RowHeight = 20
    With wsC.cells(MATRIX_ROW, MATRIX_COL + nSec + 1)
        .Value = "AVG"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter: .Font.Size = 7
    End With
    With wsC.cells(MATRIX_ROW, MATRIX_COL + nSec + 2)
        .Value = "vs SPY"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter: .Font.Size = 7
    End With
    With wsC.Range(wsC.cells(MATRIX_ROW, 1), wsC.cells(MATRIX_ROW, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    ' ---- rows ----
    Dim avgCorr() As Double: ReDim avgCorr(0 To nSec - 1)
    For i = 0 To nSec - 1
        Dim sRow As Double: sRow = 0
        For j = 0 To nSec - 1
            If i <> j Then sRow = sRow + corr(i, j)
        Next j
        avgCorr(i) = sRow / (nSec - 1)
    Next i
    Dim minAvg As Double: minAvg = 2
    Dim maxAvg As Double: maxAvg = -2
    For i = 0 To nSec - 1
        If avgCorr(i) < minAvg Then minAvg = avgCorr(i)
        If avgCorr(i) > maxAvg Then maxAvg = avgCorr(i)
    Next i

    Dim r As Long
    For i = 0 To nSec - 1
        r = MATRIX_ROW + 1 + i
        wsC.Rows(r).RowHeight = 17
        ' a thin divider whenever the group changes
        If i > 0 Then
            If groups(i) <> groups(i - 1) Then
                With wsC.Range(wsC.cells(r - 1, 1), wsC.cells(r - 1, lastCol)).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Color = RR4_LINE
                    .Weight = xlThin
                End With
            End If
        End If
        With wsC.cells(r, 1)
            .Value = labels(i)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
            .Interior.Color = RGB(10, 10, 10)
        End With
        With wsC.cells(r, 2)
            .Value = tickers(i)
            .Font.Color = RGB(140, 140, 140)
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(10, 10, 10)
        End With
        For j = 0 To nSec - 1
            With wsC.cells(r, MATRIX_COL + j)
                .HorizontalAlignment = xlCenter
                If i = j Then
                    .Value = ChrW(&H2014)
                    .Interior.Color = RGB(20, 20, 20)
                    .Font.Color = RGB(70, 70, 70)
                ElseIf Not (ok(i) And ok(j)) Then
                    .Value = "n/a"
                    .Interior.Color = RGB(14, 14, 14)
                    .Font.Color = RGB(80, 80, 80)
                Else
                    .Value = corr(i, j)
                    .NumberFormat = "0.00"
                    .Font.Bold = (Abs(corr(i, j)) >= 0.7)
                    .Interior.Color = CorrHeatBg(corr(i, j))
                    .Font.Color = CorrHeatFg(corr(i, j))
                End If
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = RGB(0, 0, 0)
                    .Weight = xlThin
                End With
            End With
        Next j
        With wsC.cells(r, MATRIX_COL + nSec + 1)
            .Value = avgCorr(i)
            .NumberFormat = "0.00"
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(12, 12, 12)
            If avgCorr(i) = minAvg Then
                .Font.Color = RGB(0, 200, 90)
            ElseIf avgCorr(i) = maxAvg Then
                .Font.Color = RGB(235, 70, 70)
            Else
                .Font.Color = RGB(200, 200, 200)
            End If
        End With
        If spyIdx >= 0 Then
            With wsC.cells(r, MATRIX_COL + nSec + 2)
                If i = spyIdx Then
                    .Value = ChrW(&H2014)
                    .Font.Color = RGB(70, 70, 70)
                Else
                    .Value = corr(i, spyIdx)
                    .NumberFormat = "0.00"
                    .Interior.Color = CorrHeatBg(corr(i, spyIdx))
                    .Font.Color = CorrHeatFg(corr(i, spyIdx))
                End If
                .HorizontalAlignment = xlCenter
            End With
        End If
    Next i
    Dim lastMatrixRow As Long: lastMatrixRow = MATRIX_ROW + nSec
    With wsC.Range(wsC.cells(lastMatrixRow, 1), wsC.cells(lastMatrixRow, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    ' ---- scale ----
    r = lastMatrixRow + 2
    wsC.cells(r, 1).Value = "SCALE"
    wsC.cells(r, 1).Font.Color = RR4_ACCENT
    wsC.cells(r, 1).Font.Bold = True
    Dim ramp As Variant: ramp = Array(-1, -0.75, -0.5, -0.25, 0, 0.25, 0.5, 0.75, 1)
    Dim k As Long
    For k = 0 To UBound(ramp)
        With wsC.cells(r, MATRIX_COL + k)
            .Value = ramp(k)
            .NumberFormat = "+0.00;-0.00;0.00"
            .HorizontalAlignment = xlCenter
            .Interior.Color = CorrHeatBg(CDbl(ramp(k)))
            .Font.Color = CorrHeatFg(CDbl(ramp(k)))
        End With
    Next k
    With wsC.cells(r, MATRIX_COL + UBound(ramp) + 2)
        .Value = "red = move together   green = move against"
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
    End With

    ' ---- group averages (within-group mean pairwise corr) ----
    r = r + 2
    wsC.cells(r, 1).Value = "WITHIN-GROUP AVG CORR"
    wsC.cells(r, 1).Font.Color = RR4_ACCENT
    wsC.cells(r, 1).Font.Bold = True
    wsC.cells(r, MATRIX_COL + 6).Value = "MOST CORRELATED PAIRS"
    wsC.cells(r, MATRIX_COL + 6).Font.Color = RR4_ACCENT
    wsC.cells(r, MATRIX_COL + 6).Font.Bold = True
    wsC.cells(r, MATRIX_COL + 15).Value = "LEAST CORRELATED / HEDGING PAIRS"
    wsC.cells(r, MATRIX_COL + 15).Font.Color = RR4_ACCENT
    wsC.cells(r, MATRIX_COL + 15).Font.Bold = True
    r = r + 1
    Dim gRow As Long: gRow = r
    Dim g As Long, seen As String
    Dim allSum As Double, allN As Long
    For g = 0 To nSec - 1
        If InStr(seen, "|" & groups(g) & "|") = 0 And groups(g) <> "BENCHMARK" Then
            seen = seen & "|" & groups(g) & "|"
            Dim gs As Double: gs = 0
            Dim gn As Long: gn = 0
            For i = 0 To nSec - 1
                For j = i + 1 To nSec - 1
                    If groups(i) = groups(g) And groups(j) = groups(g) Then gs = gs + corr(i, j): gn = gn + 1
                Next j
            Next i
            wsC.cells(gRow, 1).Value = groups(g)
            wsC.cells(gRow, 1).Font.Color = RGB(200, 200, 200)
            With wsC.cells(gRow, 2)
                If gn > 0 Then
                    .Value = gs / gn
                    .NumberFormat = "0.00"
                    .Interior.Color = CorrHeatBg(gs / gn)
                    .Font.Color = CorrHeatFg(gs / gn)
                Else
                    .Value = "n=1"
                    .Font.Color = RGB(80, 80, 80)
                End If
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
            End With
            gRow = gRow + 1
        End If
    Next g
    For i = 0 To nSec - 1
        For j = i + 1 To nSec - 1
            If ok(i) And ok(j) Then allSum = allSum + corr(i, j): allN = allN + 1
        Next j
    Next i
    wsC.cells(gRow, 1).Value = "ALL PAIRS"
    wsC.cells(gRow, 1).Font.Color = RR4_ACCENT
    wsC.cells(gRow, 1).Font.Bold = True
    With wsC.cells(gRow, 2)
        If allN > 0 Then .Value = allSum / allN
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Color = IIf(allN > 0 And allSum / allN >= 0.6, RGB(235, 70, 70), RGB(221, 221, 221))
    End With
    With wsC.cells(gRow, 3)
        .Value = IIf(allN > 0 And allSum / allN >= 0.6, "  macro-driven tape: everything moves together", _
                 IIf(allN > 0 And allSum / allN <= 0.3, "  dispersed tape: sector picking matters", "  moderate"))
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
    End With

    ' ---- top / bottom pairs (across groups only - same-group pairs are trivially high) ----
    Dim nP As Long: nP = 0
    Dim pA() As Long, pB() As Long, pV() As Double
    ReDim pA(1 To nSec * nSec): ReDim pB(1 To nSec * nSec): ReDim pV(1 To nSec * nSec)
    For i = 0 To nSec - 1
        For j = i + 1 To nSec - 1
            If ok(i) And ok(j) And groups(i) <> groups(j) And tickers(i) <> "SPY" And tickers(j) <> "SPY" Then
                nP = nP + 1: pA(nP) = i: pB(nP) = j: pV(nP) = corr(i, j)
            End If
        Next j
    Next i
    Dim a As Long, b As Long, tL As Long, tD As Double
    For a = 1 To nP - 1
        For b = a + 1 To nP
            If pV(b) > pV(a) Then
                tD = pV(a): pV(a) = pV(b): pV(b) = tD
                tL = pA(a): pA(a) = pA(b): pA(b) = tL
                tL = pB(a): pB(a) = pB(b): pB(b) = tL
            End If
        Next b
    Next a
    Dim show As Long: show = IIf(nP < 6, nP, 6)
    For k = 1 To show
        Call PairLine(wsC, r + k - 1, MATRIX_COL + 6, labels(pA(k)), labels(pB(k)), pV(k))
        Dim kk As Long: kk = nP - k + 1
        Call PairLine(wsC, r + k - 1, MATRIX_COL + 15, labels(pA(kk)), labels(pB(kk)), pV(kk))
    Next k
    With wsC.cells(r + show, MATRIX_COL + 6)
        .Value = "(cross-group pairs only, SPY excluded)"
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 8
    End With

    ' ---- notes ----
    r = IIf(gRow + 2 > r + show + 2, gRow + 2, r + show + 2)
    Dim notes As Variant
    notes = Array( _
        "HOW THESE ARE COMPUTED", _
        "CORRELATION   = Pearson correlation of simple daily returns over the last " & (nDays - 1) & " returns (Yahoo daily closes, 1y pull)", _
        "AVG CORR      = mean of an ETF's correlations with the other " & (nSec - 1) & "   - lowest (green) = the odd one out, highest (red) = the market proxy", _
        "vs SPY        = correlation with SPY   - how much of the ETF is just beta", _
        "WITHIN-GROUP  = mean pairwise correlation inside each sector-rotation category; ALL PAIRS >= 0.6 = macro-driven, <= 0.3 = dispersed", _
        "UNIVERSE      = sector-rotation-system/config.py RotationConfig.universe + SPY, grouped by its TICKER_CATEGORY")
    For k = 0 To UBound(notes)
        With wsC.cells(r + k, 1)
            .Value = notes(k)
            .Font.Size = 9
            If k = 0 Then
                .Font.Color = RR4_ACCENT: .Font.Bold = True
            Else
                .Font.Color = RGB(120, 120, 120)
            End If
        End With
    Next k

    Application.ScreenUpdating = True
    Call NavAdd(wsC, "CC")
    ' no frozen panes on this page (2026-09-12) - the FreezePanes = False at
    ' the top also clears any left over from the v1 layout
    wsC.cells(1, 1).Select
    Application.EnableEvents = prevEv
    Call NavNotify("CC done - " & nSec & " ETFs, " & nDays & " days")
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    Call NavNotify("CC failed: " & Err.Description, True)
End Sub

' "Tech x Semi   0.82" with the value in the heat colour
Private Sub PairLine(ws As Worksheet, r As Long, c As Long, t1 As String, t2 As String, v As Double)
    With ws.cells(r, c)
        .Value = t1 & " x " & t2
        .Font.Color = RGB(200, 200, 200)
    End With
    With ws.cells(r, c + 4)
        .Value = v
        .NumberFormat = "0.00"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = CorrHeatBg(v)
        .Font.Color = CorrHeatFg(v)
    End With
End Sub

' Closing price array (nDays entries, oldest first) from Yahoo's 1y daily chart.
Private Function FetchPriceArray(Ticker As String, nDays As Long) As Double()
    Dim result() As Double
    ReDim result(0 To nDays - 1)

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=1y"

    On Error Resume Next
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    On Error GoTo 0

    If http.status <> 200 Then FetchPriceArray = result: Exit Function

    Dim resp As String: resp = http.responseText
    Dim pos  As Long: pos = InStr(resp, """close"":[")
    If pos = 0 Then FetchPriceArray = result: Exit Function

    Dim arrStr As String
    arrStr = Mid(resp, pos + 9, 5000)
    arrStr = Split(arrStr, "]")(0)
    Dim parts() As String: parts = Split(arrStr, ",")

    ' Take last nDays values
    Dim total As Long: total = UBound(parts) + 1
    Dim startIdx As Long: startIdx = IIf(total > nDays, total - nDays, 0)
    Dim idx As Long: idx = 0
    Dim k As Long
    For k = startIdx To UBound(parts)
        If IsNumeric(Trim(parts(k))) And idx < nDays Then
            result(idx) = CDbl(Trim(parts(k)))
            idx = idx + 1
        End If
    Next k

    FetchPriceArray = result
End Function

' Pearson correlation between two return series (rows of retData)
Private Function CalcCorrelation(retData() As Double, r1 As Long, r2 As Long, n As Long) As Double
    Dim sumX As Double, sumY As Double, sumXY As Double
    Dim sumX2 As Double, sumY2 As Double
    Dim i As Long, cnt As Long

    For i = 0 To n - 1
        Dim x As Double: x = retData(r1, i)
        Dim y As Double: y = retData(r2, i)
        If x <> 0 Or y <> 0 Then
            sumX = sumX + x
            sumY = sumY + y
            sumXY = sumXY + x * y
            sumX2 = sumX2 + x * x
            sumY2 = sumY2 + y * y
            cnt = cnt + 1
        End If
    Next i

    If cnt < 5 Then CalcCorrelation = 0: Exit Function

    Dim num As Double: num = cnt * sumXY - sumX * sumY
    Dim den As Double: den = Sqr((cnt * sumX2 - sumX ^ 2) * (cnt * sumY2 - sumY ^ 2))
    If den = 0 Then CalcCorrelation = 0 Else CalcCorrelation = num / den
End Function

Sub RefreshCorrelation()
    Call BuildCorrelationMatrix
End Sub
