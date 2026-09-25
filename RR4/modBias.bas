Attribute VB_Name = "modBias"
Option Explicit

' ================================================================
'  BIAS (nav code B, B! = rebuild) - sheet "Bias"   2026-09-23, v2 2026-09-24
' ----------------------------------------------------------------
'  User's own normalized-EMA-deviation formula (pasted in chat
'  2026-09-23), reduced from 4 lines to 2 per the user's instruction:
'  keep only the shortest (N1=20) and longest (N4=200) line.
'    B1 = (C - EMA(C,N1)) / EMA(C,N1) * 100
'    B4 = (C - EMA(C,N4)) / EMA(C,N4) * 100
'    R1 = B1>=0 ? B1/Max(HHV(B1,M),0.01)*100 : B1/Max(Abs(LLV(B1,M)),0.01)*100
'    R4 = same shape, from B4
'  N1=20 N4=200 M=250 K=0.8. The original 4-line cross-count signal
'  (UPCNT>=3 / DNCNT>=2) has no meaning left with only 2 lines; see the
'  2026-09-23 second-round note below for what replaced it. The
'  HOLDINGS/WATCHLIST tables below remain colour-intensity only, no
'  signal column (see CorrHeatBg/Fg reuse) - only the chart has signals.
'
'  Universe: HOLDINGS (the RR4 position table) and WATCHLIST are two
'  SEPARATE tables (2026-09-24, user's request), each deduped and sorted
'  by R1 on its own - NOT a full market scan (a full TW+US sweep is
'  RRG-Industry-scale, tens of minutes to hours, wrong tool for an
'  on-demand page). A separate TICKER query box draws one ticker's full
'  time series (all 5 reference lines) on demand; that ticker is never
'  folded into either table.
'
'  History: modvolatility.GetHistoricalData (Yahoo chart API, already
'  handles a bare TW numeric code by trying .TW then .TWO), trimmed to
'  the newest NEED_BARS closes (~2y, per the user: "just enough" for the
'  N4=200 EMA warmup plus the M=250 normalization window).
'  R1 (short) and R4 (long) each have their OWN minimum-history floor
'  (N+M_WIN) - a ticker with too little history for R4 (e.g. a recent
'  IPO) still gets R1 computed and shown; only R4 shows "-" for that row
'  (2026-09-24, replaces the earlier all-or-nothing "n/a" row).
'
'  B! is the only thing that (re)fetches and rebuilds the tables; jumping
'  to the page does not, same convention as every other report page
'  (V!/HC!/RGE! etc). Typing a ticker into the query box redraws only the
'  chart - it does not touch the tables or re-fetch them.
'
'  2026-09-23, second round (same day, later chat): user reported the
'  ORIGINAL 4-line UPCNT/DNCNT design throws too many false signals.
'  Redesigned discrete signals for R1/R4 - each independently, NO
'  cross-line vote, NO multi-bar persistence filter (user explicitly
'  rejected both when offered) - via two changes the user picked after
'  a design discussion:
'    C - trend filter: EMA(C,N4) slope, now vs TFLEN=10 bars ago, gates
'        each direction. A "topped out" (SIGUP) signal is suppressed
'        while N4 is still rising; a "bottomed out" (SIGDN) signal is
'        suppressed while N4 is still falling. Targets the "reversal
'        signal fires against an established trend" failure mode.
'    D - NEW lines only: NormOne's HHV/LLV min-max is replaced by
'        NormRank, a signed rolling percentile rank over MR_RANK=200
'        bars (user's own cap, must not exceed 200). One outlier bar can
'        no longer set the ceiling/floor for the whole window; K_THRESH
'        =0.8 now reads as "today beats 90% of the last 200 days".
'  Per the user's explicit request this did NOT initially replace the OLD
'  HHV/LLV lines - both were drawn on the same per-ticker chart (old thin/
'  muted, new thick) so they could be eyeballed side by side.
'
'  2026-09-23, third round (same day): after comparing both on the chart,
'  the user chose the NEW (rank+trend) line only - the OLD HHV/LLV line is
'  REMOVED from the chart (ComputeOneLine/NormOne are UNCHANGED and still
'  drive the HOLDINGS/WATCHLIST tables, which were never part of this
'  comparison and are untouched). RightAlignDbl/RightAlignBool, which only
'  existed to line up the old and new series on one date axis, are gone
'  too - the chart is now driven end to end by ComputeOneLineRankSignal.
'  R1/R4 line weight simplified to 1pt now that there is only one line per
'  side (no longer needs to out-weigh a muted comparison line).
'
'  2026-09-24, K-line: each R chart now has a K-line chart + volume chart
'  to its RIGHT over the same bars (R1 SHORT pair: candles + EMA20 + short
'  SELL/BUY triangles; R4 LONG pair: candles + EMA200 + long triangles).
'  OHLCV is its own Yahoo fetch (adjusted O/H/L, see FetchBiasOhlcvRaw),
'  aligned by date to the chart bars. See the DrawBiasKlines block for why
'  EMA/triangles sit in a secondary axis group.
'
'  Pure ASCII (VBE import rule).
' ================================================================

Public Const BIAS_SHEET As String = "Bias"

Private Const N1 As Long = 20
Private Const N4 As Long = 200
Private Const M_WIN As Long = 250
Private Const K_THRESH As Double = 0.8
Private Const NEED_BARS As Long = 500
Private Const MIN_BARS_SHORT As Long = N1 + M_WIN   ' 270 - floor for R1 alone (fetch-worthy floor)

' --- 2026-09-23 second round: NEW rank+trend line/signal, chart-only ---
Private Const MR_RANK As Long = 200   ' D: percentile-rank window, user's cap (must not exceed 200)
Private Const TFLEN As Long = 10      ' C: trend-filter lookback for EMA(N4) slope

Private Const PG_TITLE As Long = 1
Private Const PG_LBL As Long = 2
Private Const PG_IN As Long = 3
Private Const COL_TK As Long = 1

Private Const HOLD_TITLE_ROW As Long = 5    ' HOLDINGS table title row; WATCHLIST follows dynamically
Private Const CHART_ROW As Long = 5         ' the ticker chart anchors at this row, in its own column block
Private Const HEAT_COL_TICKER As Long = 1
Private Const HEAT_COL_MKT As Long = 2
Private Const HEAT_COL_LAST As Long = 3
Private Const HEAT_COL_CHG As Long = 4
Private Const HEAT_COL_R1 As Long = 5
Private Const HEAT_COL_R4 As Long = 6

Private Const CHART_COL As Long = 8          ' physical column I after the bar's +1 column shift
Private Const CHART_DATA_COL As Long = 676   ' hidden data block the chart series point at: physical column ZA (676 + the bar's +1 column shift = 677)
Private Const LEGACY_DATA_COL As Long = 30   ' where the block lived before 2026-09-25 (physical AE); cleared on every rebuild
Private Const BIAS_CHART_H As Double = 460      ' each of the two stacked charts
Private Const BIAS_CHART_W As Double = 760
Private Const BIAS_CHART_GAP As Double = 10

' --- 2026-09-24: K-line + volume charts, right of each R chart ---
Private Const KLINE_GAP As Double = 12          ' horizontal gap between an R chart and its K-line chart
Private Const KLINE_H As Double = 335           ' K-line chart height; KLINE_H + KV_GAP + VOL_H = BIAS_CHART_H
Private Const VOL_H As Double = 120
Private Const KV_GAP As Double = 5
Private Const KB_OFF As Long = 13               ' K/volume data block starts this many columns after CHART_DATA_COL
Private Const KB_LAST As Long = 24              ' last data column offset (KB_OFF .. KB_LAST = 12 columns)
Private Const K_PLOT_LEFT As Double = 58        ' plot-area inside left/right margins, shared so K and volume line up
Private Const K_PLOT_RIGHT As Double = 12

Private Const FONT_FACE As String = "Consolas"
Private Const CLR_TEXT As Long = 14540253            ' RGB(221,221,221)
Private Const CLR_MUTED As Long = 8553090            ' RGB(130,130,130)

' ----------------------------------------------------------------
'  Entry points
' ----------------------------------------------------------------

' B! - rebuild both tables (holdings, watchlist) and, if a ticker is
' already sitting in the query box, redraw its chart too.
Public Sub BuildBiasPage()
    Dim ws As Worksheet: Set ws = EnsureBiasSheet()
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fail

    Dim keepTk As String
    keepTk = UCase(Trim(CellStr(ws.cells(PG_IN + NavOffset(ws), COL_TK + NavLeft(ws)).Value)))

    Call NavStrip(ws)
    Call DrawShell(ws, keepTk)
    Call ClearBiasChart(ws)             ' drop any stale chart from a previous build up front
    Call NavNotify("BIAS building tables (holdings + watchlist) ...")

    Dim posT() As String, wlT() As String
    Dim nPos As Long, nWl As Long
    nPos = GatherPositions(posT)
    nWl = GatherWatchlist(wlT)

    Dim tk1() As String, mkt1() As String, last1() As Double, chg1() As Double
    Dim r1a() As Double, r4a() As Double, hasR4a() As Boolean, ok1() As Boolean
    Dim tk2() As String, mkt2() As String, last2() As Double, chg2() As Double
    Dim r2a() As Double, r5a() As Double, hasR5a() As Boolean, ok2() As Boolean
    Dim doneCount As Long

    Call BuildTableData(posT, nPos, "HOLDINGS", tk1, mkt1, last1, chg1, r1a, hasR4a, r4a, ok1, doneCount)
    Call BuildTableData(wlT, nWl, "WATCHLIST", tk2, mkt2, last2, chg2, r2a, hasR5a, r5a, ok2, doneCount)

    Dim nextRow As Long: nextRow = HOLD_TITLE_ROW
    nextRow = DrawOneTable(ws, "HOLDINGS", tk1, mkt1, last1, chg1, r1a, hasR4a, r4a, ok1, nPos, nextRow)
    nextRow = DrawOneTable(ws, "WATCHLIST", tk2, mkt2, last2, chg2, r2a, hasR5a, r5a, ok2, nWl, nextRow)

    Call NavNotify("BIAS done - " & doneCount & "/" & (nPos + nWl) & " tickers")

    If Len(keepTk) > 0 Then Call RefreshBiasQuery

    Call FinishPage(ws)
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
    Call NavGoto("B", ws)
    Exit Sub

Fail:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    Call NavNotify("BIAS error: " & msg, True)
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
End Sub

' Sheet event (Worksheet_Change, written by EnsureSheetCode): typing a
' ticker in the query box redraws only its chart, never the tables.
Public Sub BiasChange(ByVal ws As Worksheet, ByVal Target As Range)
    If Target.CountLarge > 4 Then Exit Sub
    Dim r As Long: r = PG_IN + NavOffset(ws)
    Dim c As Long: c = COL_TK + NavLeft(ws)
    If Intersect(Target, ws.cells(r, c)) Is Nothing Then Exit Sub
    Call RefreshBiasQuery
End Sub

Public Sub RefreshBiasQuery()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(BIAS_SHEET): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim tk As String: tk = UCase(Trim(CellStr(ws.cells(PG_IN + NavOffset(ws), COL_TK + NavLeft(ws)).Value)))
    Call ClearBiasChart(ws)
    If tk = "" Then Exit Sub

    Call NavNotify("BIAS fetching " & tk & " ...")
    Dim closeArr() As Double, dateArr() As Date, cnt As Long
    If Not FetchBiasHistory(tk, closeArr, dateArr, cnt) Then
        Call NavNotify("BIAS " & tk & ": no price history", True)
        Exit Sub
    End If

    ' 2026-09-23 third round: chart is driven end to end by the NEW
    ' rank+trend line/signal - see ComputeOneLineRankSignal header note.
    ' Its floor is shared between R1 and R4 (both gated on N4+MR_RANK, the
    ' trend filter's own requirement), so vc1 and vc4 are always equal
    ' whenever hasR4 is True - no alignment step needed between them.
    Dim rr1() As Double, sigUp1() As Boolean, sigDn1() As Boolean, vc1 As Long
    If Not ComputeOneLineRankSignal(closeArr, cnt, N1, rr1, sigUp1, sigDn1, vc1) Or vc1 = 0 Then
        Call NavNotify("BIAS " & tk & ": not enough history for the rank+trend window", True)
        Exit Sub
    End If
    Dim rr4() As Double, sigUp4() As Boolean, sigDn4() As Boolean, vc4 As Long, hasR4 As Boolean
    hasR4 = ComputeOneLineRankSignal(closeArr, cnt, N4, rr4, sigUp4, sigDn4, vc4) And vc4 > 0

    Dim chartDates() As Date, i As Long
    ReDim chartDates(0 To vc1 - 1)
    For i = 0 To vc1 - 1
        chartDates(i) = dateArr(cnt - vc1 + i)
    Next i

    Call DrawBiasChart(ws, tk, chartDates, rr1, hasR4, rr4, sigUp1, sigDn1, sigUp4, sigDn4, vc1)

    ' 2026-09-24: K-line + volume beside each R chart, same date span (the
    ' vc1 newest bars). OHLCV is a separate Yahoo fetch aligned by date to
    ' chartDates; EMAs are recomputed here from the same closes the R lines use.
    Dim ko() As Double, kh() As Double, kl() As Double, kc() As Double, kv() As Double
    Dim kOk As Boolean
    kOk = FetchBiasOhlc(tk, chartDates, vc1, closeArr, cnt, ko, kh, kl, kc, kv)
    If kOk Then
        Dim bTmp() As Double, emaAllS() As Double, emaAllL() As Double
        Call ComputeEmaDeviation(closeArr, cnt, N1, bTmp, emaAllS)
        Call ComputeEmaDeviation(closeArr, cnt, N4, bTmp, emaAllL)
        Dim emaS() As Double, emaL() As Double
        ReDim emaS(0 To vc1 - 1): ReDim emaL(0 To vc1 - 1)
        For i = 0 To vc1 - 1
            emaS(i) = emaAllS(cnt - vc1 + i)
            emaL(i) = emaAllL(cnt - vc1 + i)
        Next i
        Call DrawBiasKlines(ws, tk, vc1, ko, kh, kl, kc, kv, emaS, emaL, hasR4, sigUp1, sigDn1, sigUp4, sigDn4)
    End If
    Call NavNotify("BIAS " & tk & " done - " & vc1 & " points" & IIf(hasR4, "", " (short line only)") & _
                   IIf(kOk, "", " (K-line unavailable: no OHLC)"))
End Sub

' ----------------------------------------------------------------
'  Universe: RR4 positions and WATCHLIST, each deduped separately
' ----------------------------------------------------------------
Private Function GatherPositions(ByRef outTickers() As String) As Long
    Dim wsP As Worksheet
    On Error Resume Next: Set wsP = ThisWorkbook.Worksheets(NavSheetName("P")): On Error GoTo 0
    If wsP Is Nothing Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim list As New Collection
    Dim r As Long: r = RR4_POS_FIRST
    Do While CellStr(wsP.cells(r, RR4_LEFT + 1).Value) <> "" And r < RR4_POS_FIRST + 2000
        Dim tk As String: tk = UCase(Trim(CellStr(wsP.cells(r, RR4_LEFT + 1).Value)))
        If tk <> "" Then
            If Not seen.Exists(tk) Then
                seen(tk) = True
                list.Add tk
            End If
        End If
        r = r + 1
    Loop

    Dim n As Long: n = list.count
    If n = 0 Then Exit Function
    ReDim outTickers(0 To n - 1)
    Dim i As Long
    For i = 1 To n
        outTickers(i - 1) = list(i)
    Next i
    GatherPositions = n
End Function

Private Function GatherWatchlist(ByRef outTickers() As String) As Long
    Dim wsP As Worksheet
    On Error Resume Next: Set wsP = ThisWorkbook.Worksheets(NavSheetName("P")): On Error GoTo 0
    If wsP Is Nothing Then Exit Function
    If Left(UCase(CellStr(wsP.cells(RR4_WL_TITLE, RR4_LEFT + 1).Value)), 9) <> "WATCHLIST" Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim list As New Collection
    Dim wr As Long
    For wr = RR4_WL_FIRST To RR4_WL_LAST
        Dim wtk As String: wtk = UCase(Trim(CellStr(wsP.cells(wr, RR4_LEFT + 1).Value)))
        If wtk <> "" Then
            If Not seen.Exists(wtk) Then
                seen(wtk) = True
                list.Add wtk
            End If
        End If
    Next wr

    Dim n As Long: n = list.count
    If n = 0 Then Exit Function
    ReDim outTickers(0 To n - 1)
    Dim i As Long
    For i = 1 To n
        outTickers(i - 1) = list(i)
    Next i
    GatherWatchlist = n
End Function

' Fetch + compute + sort one ticker list into its parallel result arrays
' (shared by the HOLDINGS and WATCHLIST tables so the logic exists once).
Private Sub BuildTableData(ByRef srcTickers() As String, ByVal n As Long, ByVal label As String, _
                            ByRef outTk() As String, ByRef outMkt() As String, ByRef outLast() As Double, _
                            ByRef outChg() As Double, ByRef outR1() As Double, ByRef outHasR4() As Boolean, _
                            ByRef outR4() As Double, ByRef outOk() As Boolean, ByRef doneCount As Long)
    If n = 0 Then Exit Sub
    ReDim outTk(0 To n - 1): ReDim outMkt(0 To n - 1): ReDim outLast(0 To n - 1)
    ReDim outChg(0 To n - 1): ReDim outR1(0 To n - 1): ReDim outHasR4(0 To n - 1)
    ReDim outR4(0 To n - 1): ReDim outOk(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        outTk(i) = srcTickers(i)
        outMkt(i) = IIf(InStr(outTk(i), ".TW") > 0, "TW", "US")
        Call NavNotify("BIAS " & label & " " & (i + 1) & "/" & n & " " & outTk(i) & " ...")

        Dim lastV As Double, chgV As Double, r1V As Double, r4V As Double, hasR4V As Boolean
        If FetchAndCompute(outTk(i), lastV, chgV, r1V, hasR4V, r4V) Then
            outLast(i) = lastV: outChg(i) = chgV: outR1(i) = r1V: outHasR4(i) = hasR4V: outR4(i) = r4V
            outOk(i) = True
            doneCount = doneCount + 1
        End If
    Next i
    Call SortByR1(outTk, outMkt, outLast, outChg, outR1, outR4, outHasR4, outOk, n)
End Sub

' ----------------------------------------------------------------
'  Price history + the EMA-deviation math
' ----------------------------------------------------------------

' Wraps modvolatility.GetHistoricalData and trims to the newest NEED_BARS
' closes (oldest -> newest). Floor is MIN_BARS_SHORT (R1's own floor) so
' a ticker with too little history for R4 (e.g. a recent IPO) still gets
' fetched and gets R1 - see ComputeOneLine for the per-line floor.
Private Function FetchBiasHistory(ByVal ticker As String, ByRef outClose() As Double, _
                                   ByRef outDate() As Date, ByRef outCount As Long) As Boolean
    Dim dates() As Date, prices() As Double
    If Not modvolatility.GetHistoricalData(ticker, dates, prices) Then Exit Function
    Dim n As Long: n = UBound(prices) - LBound(prices) + 1
    If n < MIN_BARS_SHORT Then Exit Function

    Dim keep As Long: keep = n: If keep > NEED_BARS Then keep = NEED_BARS
    Dim startIdx As Long: startIdx = LBound(prices) + (n - keep)
    ReDim outClose(0 To keep - 1): ReDim outDate(0 To keep - 1)
    Dim i As Long
    For i = 0 To keep - 1
        outClose(i) = prices(startIdx + i)
        outDate(i) = dates(startIdx + i)
    Next i
    outCount = keep
    FetchBiasHistory = True
End Function

' One ticker, both lines: R1 is required (a fetched ticker always clears
' MIN_BARS_SHORT); R4 is optional and simply omitted (outHasR4 = False)
' when there is not yet enough history for EMA200's own floor (N4+M_WIN).
Private Function FetchAndCompute(ByVal ticker As String, ByRef outLast As Double, ByRef outChg As Double, _
                                  ByRef outR1 As Double, ByRef outHasR4 As Boolean, ByRef outR4 As Double) As Boolean
    outLast = 0: outChg = 0: outR1 = 0: outHasR4 = False: outR4 = 0
    Dim closeArr() As Double, dateArr() As Date, cnt As Long
    If Not FetchBiasHistory(ticker, closeArr, dateArr, cnt) Then Exit Function

    Dim rr1() As Double, vc1 As Long
    If Not ComputeOneLine(closeArr, cnt, N1, rr1, vc1) Or vc1 = 0 Then Exit Function
    outR1 = rr1(vc1 - 1)
    outLast = closeArr(cnt - 1)
    If cnt >= 2 Then
        Dim prevC As Double: prevC = closeArr(cnt - 2)
        If prevC <> 0 Then outChg = (closeArr(cnt - 1) / prevC - 1) * 100
    End If

    Dim rr4() As Double, vc4 As Long
    If ComputeOneLine(closeArr, cnt, N4, rr4, vc4) And vc4 > 0 Then
        outHasR4 = True
        outR4 = rr4(vc4 - 1)
    End If

    FetchAndCompute = True
End Function

' Shared EMA(linePeriod) + %-deviation-from-EMA series (the "B" series in
' the user's formula). Both the OLD HHV/LLV normalization and the NEW
' percentile-rank normalization (D) start from this same B-series; only
' the final per-window normalization step differs between them. outEma is
' also returned since the NEW trend filter (C) needs the raw EMA(N4), not
' just its deviation.
Private Sub ComputeEmaDeviation(ByRef c() As Double, ByVal cnt As Long, ByVal linePeriod As Long, _
                                 ByRef outB() As Double, ByRef outEma() As Double)
    ReDim outEma(0 To cnt - 1): ReDim outB(0 To cnt - 1)
    Dim alpha As Double: alpha = 2# / (linePeriod + 1)
    outEma(0) = c(0)
    Dim i As Long
    For i = 1 To cnt - 1
        outEma(i) = alpha * c(i) + (1 - alpha) * outEma(i - 1)
    Next i
    For i = 0 To cnt - 1
        If outEma(i) <> 0 Then outB(i) = (c(i) - outEma(i)) / outEma(i) * 100
    Next i
End Sub

' EMA(N) deviation normalized by its own rolling M_WIN HHV/LLV - one line
' of the user's ORIGINAL formula, kept byte-for-byte unchanged (2026-09-23
' second round: shown alongside the new rank-based line for side-by-side
' comparison, see header comment - this function still drives the
' HOLDINGS/WATCHLIST tables untouched). Requires cnt >= N+M_WIN (EMA
' warmup floor); outR is 0-based, index 0 = the oldest point with a full
' M_WIN window of B-values behind it.
Private Function ComputeOneLine(ByRef c() As Double, ByVal cnt As Long, ByVal linePeriod As Long, _
                                 ByRef outR() As Double, ByRef outCount As Long) As Boolean
    outCount = 0
    If cnt < linePeriod + M_WIN Then Exit Function

    Dim b() As Double, ema() As Double
    Call ComputeEmaDeviation(c, cnt, linePeriod, b, ema)

    Dim n As Long: n = cnt - M_WIN + 1
    If n <= 0 Then Exit Function
    ReDim outR(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        Dim atIdx As Long: atIdx = i + M_WIN - 1
        outR(i) = NormOne(b, atIdx, atIdx - M_WIN + 1)
    Next i
    outCount = n
    ComputeOneLine = True
End Function

' R = v>=0 ? v/Max(HHV(window),0.01)*100 : v/Max(Abs(LLV(window)),0.01)*100
Private Function NormOne(ByRef b() As Double, ByVal atIdx As Long, ByVal winStart As Long) As Double
    Dim hhv As Double, llv As Double, j As Long
    hhv = b(winStart): llv = b(winStart)
    For j = winStart + 1 To atIdx
        If b(j) > hhv Then hhv = b(j)
        If b(j) < llv Then llv = b(j)
    Next j
    Dim v As Double: v = b(atIdx)
    If v >= 0 Then
        Dim denomH As Double: denomH = hhv: If denomH < 0.01 Then denomH = 0.01
        NormOne = v / denomH * 100
    Else
        Dim denomL As Double: denomL = Abs(llv): If denomL < 0.01 Then denomL = 0.01
        NormOne = v / denomL * 100
    End If
End Function

' D (2026-09-23 second round): signed percentile rank of b(atIdx) within
' the trailing MR_RANK-bar window, rescaled to -100..100 - replaces
' NormOne for the NEW line only (see header comment). 0 = today sits at
' the window's median, +100 = today is the window's highest value, -100 =
' today is the window's lowest. Unlike NormOne, one outlier bar cannot set
' the ceiling/floor for every other day in the window - it can only ever
' contribute its own +/-1 to the count.
Private Function NormRank(ByRef b() As Double, ByVal atIdx As Long, ByVal winStart As Long) As Double
    Dim todayV As Double: todayV = b(atIdx)
    Dim n As Long: n = atIdx - winStart + 1
    Dim hits As Long, j As Long
    For j = winStart To atIdx
        If b(j) <= todayV Then hits = hits + 1
    Next j
    NormRank = (hits / n * 100 - 50) * 2
End Function

' NEW (2026-09-23 second round) percentile-rank line (D) + discrete signal
' (C: EMA(N4) trend filter, single-bar edge trigger, NO persistence, NO
' cross-line vote - all three explicitly rejected by the user, see header
' comment). One call per line (linePeriod = N1 or N4); the trend filter
' always uses EMA(N4) regardless of which line is being signalled.
' outR/outSigUp/outSigDn are 0-based and aligned 1:1 (index i is the same
' calendar bar in each), anchored so the LAST element is always "today"
' (c(cnt-1)). The floor check below is the SAME for every linePeriod (it
' is really the trend filter's N4+MR_RANK requirement, not the line's own
' N+MR_RANK), which is what guarantees the R1 and R4 calls always agree on
' outCount when both succeed - no alignment step needed by the caller.
Private Function ComputeOneLineRankSignal(ByRef c() As Double, ByVal cnt As Long, ByVal linePeriod As Long, _
                                           ByRef outR() As Double, ByRef outSigUp() As Boolean, _
                                           ByRef outSigDn() As Boolean, ByRef outCount As Long) As Boolean
    outCount = 0
    If cnt < linePeriod + MR_RANK Then Exit Function
    If cnt < N4 + MR_RANK Then Exit Function   ' trend filter needs EMA(N4) warmed up over the same span

    Dim b() As Double, emaTmp() As Double, bN4() As Double, emaN4() As Double
    Call ComputeEmaDeviation(c, cnt, linePeriod, b, emaTmp)
    Call ComputeEmaDeviation(c, cnt, N4, bN4, emaN4)

    Dim n As Long: n = cnt - MR_RANK + 1
    If n <= 0 Then Exit Function
    ReDim outR(0 To n - 1): ReDim outSigUp(0 To n - 1): ReDim outSigDn(0 To n - 1)

    Dim i As Long, prevUp As Boolean, prevDn As Boolean
    For i = 0 To n - 1
        Dim atIdx As Long: atIdx = i + MR_RANK - 1
        outR(i) = NormRank(b, atIdx, atIdx - MR_RANK + 1)

        Dim up As Boolean: up = outR(i) >= K_THRESH * 100
        Dim dn As Boolean: dn = outR(i) <= -K_THRESH * 100
        If i > 0 And atIdx >= TFLEN Then
            Dim trendUp As Boolean: trendUp = emaN4(atIdx) > emaN4(atIdx - TFLEN)
            Dim trendDn As Boolean: trendDn = emaN4(atIdx) < emaN4(atIdx - TFLEN)
            outSigUp(i) = up And Not prevUp And Not trendUp
            outSigDn(i) = dn And Not prevDn And Not trendDn
        End If
        prevUp = up: prevDn = dn
    Next i
    outCount = n
    ComputeOneLineRankSignal = True
End Function

Private Sub SortByR1(ByRef tk() As String, ByRef mkt() As String, ByRef lastPx() As Double, _
                      ByRef chgPct() As Double, ByRef r1() As Double, ByRef r4() As Double, _
                      ByRef hasR4() As Boolean, ByRef okF() As Boolean, ByVal n As Long)
    Dim i As Long, j As Long, best As Long
    For i = 0 To n - 2
        best = i
        For j = i + 1 To n - 1
            If (okF(j) And Not okF(best)) Or (okF(j) = okF(best) And r1(j) > r1(best)) Then best = j
        Next j
        If best <> i Then
            Call SwapStr(tk, i, best): Call SwapStr(mkt, i, best)
            Call SwapDbl(lastPx, i, best): Call SwapDbl(chgPct, i, best)
            Call SwapDbl(r1, i, best): Call SwapDbl(r4, i, best)
            Call SwapBool(hasR4, i, best): Call SwapBool(okF, i, best)
        End If
    Next i
End Sub

Private Sub SwapStr(ByRef a() As String, ByVal i As Long, ByVal j As Long)
    Dim t As String: t = a(i): a(i) = a(j): a(j) = t
End Sub
Private Sub SwapDbl(ByRef a() As Double, ByVal i As Long, ByVal j As Long)
    Dim t As Double: t = a(i): a(i) = a(j): a(j) = t
End Sub
Private Sub SwapBool(ByRef a() As Boolean, ByVal i As Long, ByVal j As Long)
    Dim t As Boolean: t = a(i): a(i) = a(j): a(j) = t
End Sub

' ----------------------------------------------------------------
'  One table (HOLDINGS or WATCHLIST) - returns the next free title row
'  so the caller can stack a second table right after this one.
' ----------------------------------------------------------------
Private Function DrawOneTable(ByVal ws As Worksheet, ByVal title As String, ByRef tk() As String, _
                               ByRef mkt() As String, ByRef lastPx() As Double, ByRef chgPct() As Double, _
                               ByRef r1() As Double, ByRef hasR4() As Boolean, ByRef r4() As Double, _
                               ByRef okF() As Boolean, ByVal n As Long, ByVal titleRow As Long) As Long
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim hdrRow As Long: hdrRow = titleRow + 1
    Dim firstRow As Long: firstRow = titleRow + 2

    With ws.cells(titleRow + off, HEAT_COL_TICKER + lc)
        .Value = UCase(title)
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 10
    End With

    Dim hdr As Variant: hdr = Array("TICKER", "MKT", "LAST", "CHG%", "R1 (EMA" & N1 & ")", "R4 (EMA" & N4 & ")")
    Dim c As Long
    For c = 0 To 5
        With ws.cells(hdrRow + off, HEAT_COL_TICKER + c + lc)
            .Value = hdr(c)
            .Font.Color = RGB(0, 200, 255)
            .Font.Size = 9
            .HorizontalAlignment = IIf(c = 0, xlLeft, xlCenter)
        End With
    Next c
    With ws.Range(ws.cells(hdrRow + off, HEAT_COL_TICKER + lc), ws.cells(hdrRow + off, HEAT_COL_R4 + lc)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With

    If n = 0 Then
        With ws.cells(firstRow + off, HEAT_COL_TICKER + lc)
            .Value = "(none)"
            .Font.Color = CLR_MUTED
        End With
        DrawOneTable = firstRow + 2
        Exit Function
    End If

    Dim i As Long
    For i = 0 To n - 1
        Dim r As Long: r = firstRow + off + i
        ws.cells(r, HEAT_COL_TICKER + lc).Value = tk(i)
        ws.cells(r, HEAT_COL_TICKER + lc).Font.Color = CLR_TEXT
        ws.cells(r, HEAT_COL_MKT + lc).Value = mkt(i)
        ws.cells(r, HEAT_COL_MKT + lc).Font.Color = CLR_MUTED
        ws.cells(r, HEAT_COL_MKT + lc).HorizontalAlignment = xlCenter

        If okF(i) Then
            With ws.cells(r, HEAT_COL_LAST + lc)
                .Value = lastPx(i): .NumberFormat = "#,##0.00": .Font.Color = CLR_TEXT: .HorizontalAlignment = xlCenter
            End With
            With ws.cells(r, HEAT_COL_CHG + lc)
                .Value = chgPct(i): .NumberFormat = "+0.00;-0.00"
                .Font.Color = IIf(chgPct(i) >= 0, RGB(220, 60, 60), RGB(60, 160, 90))
                .HorizontalAlignment = xlCenter
            End With
            With ws.cells(r, HEAT_COL_R1 + lc)
                .Value = r1(i): .NumberFormat = "0.0"
                .Interior.Color = CorrHeatBg(r1(i) / 100): .Font.Color = CorrHeatFg(r1(i) / 100)
                .Font.Bold = True: .HorizontalAlignment = xlCenter
            End With
            With ws.cells(r, HEAT_COL_R4 + lc)
                If hasR4(i) Then
                    .Value = r4(i): .NumberFormat = "0.0"
                    .Interior.Color = CorrHeatBg(r4(i) / 100): .Font.Color = CorrHeatFg(r4(i) / 100)
                    .Font.Bold = True
                Else
                    .Value = "-"
                    .Interior.Color = RGB(0, 0, 0)
                    .Font.Color = CLR_MUTED
                    .Font.Bold = False
                End If
                .HorizontalAlignment = xlCenter
            End With
        Else
            ws.cells(r, HEAT_COL_LAST + lc).Value = "n/a"
            ws.Range(ws.cells(r, HEAT_COL_LAST + lc), ws.cells(r, HEAT_COL_R4 + lc)).Font.Color = CLR_MUTED
        End If
    Next i

    Dim lastR As Long: lastR = firstRow + off + n - 1
    With ws.Range(ws.cells(lastR, HEAT_COL_TICKER + lc), ws.cells(lastR, HEAT_COL_R4 + lc)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With

    DrawOneTable = firstRow + n + 1   ' one blank row of gap before the next table's title
End Function

' ----------------------------------------------------------------
'  Single-ticker time series chart (all 5 reference lines)
' ----------------------------------------------------------------
' NOTE (2026-09-23 third round): r1()/r4() are the rank+trend line (the
' OLD HHV/LLV line was removed from the chart this round, see header
' comment); r4()/sigUp4()/sigDn4() are only meaningful when hasR4 is True.
Private Sub DrawBiasChart(ByVal ws As Worksheet, ByVal ticker As String, ByRef d() As Date, _
                           ByRef r1() As Double, ByVal hasR4 As Boolean, ByRef r4() As Double, _
                           ByRef sigUp1() As Boolean, ByRef sigDn1() As Boolean, _
                           ByRef sigUp4() As Boolean, ByRef sigDn4() As Boolean, ByVal n As Long)
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim dc As Long: dc = CHART_DATA_COL + lc

    ws.cells(CHART_ROW + off, dc).Value = "date"
    ws.cells(CHART_ROW + off, dc + 1).Value = "R1"
    ws.cells(CHART_ROW + off, dc + 2).Value = "R4"
    ws.cells(CHART_ROW + off, dc + 3).Value = "TOP"
    ws.cells(CHART_ROW + off, dc + 4).Value = "UPPER"
    ws.cells(CHART_ROW + off, dc + 5).Value = "ZERO"
    ws.cells(CHART_ROW + off, dc + 6).Value = "LOWER"
    ws.cells(CHART_ROW + off, dc + 7).Value = "BOTTOM"
    ws.cells(CHART_ROW + off, dc + 8).Value = "SIGUP1"
    ws.cells(CHART_ROW + off, dc + 9).Value = "SIGDN1"
    ws.cells(CHART_ROW + off, dc + 10).Value = "SIGUP4"
    ws.cells(CHART_ROW + off, dc + 11).Value = "SIGDN4"
    ws.Range(ws.cells(CHART_ROW + off, dc), ws.cells(CHART_ROW + off, dc + 11)).Font.Color = RGB(90, 90, 90)

    Dim i As Long, r As Long
    For i = 0 To n - 1
        r = CHART_ROW + off + 1 + i
        ws.cells(r, dc).Value = d(i): ws.cells(r, dc).NumberFormat = "yyyy-mm-dd"
        ws.cells(r, dc + 1).Value = r1(i)
        If hasR4 Then ws.cells(r, dc + 2).Value = r4(i)
        ws.cells(r, dc + 3).Value = 100
        ws.cells(r, dc + 4).Value = K_THRESH * 100
        ws.cells(r, dc + 5).Value = 0
        ws.cells(r, dc + 6).Value = -K_THRESH * 100
        ws.cells(r, dc + 7).Value = -100
        If sigUp1(i) Then ws.cells(r, dc + 8).Value = 112
        If sigDn1(i) Then ws.cells(r, dc + 9).Value = -112
        If hasR4 Then
            If sigUp4(i) Then ws.cells(r, dc + 10).Value = 124
            If sigDn4(i) Then ws.cells(r, dc + 11).Value = -124
        End If
    Next i
    ws.Range(ws.cells(CHART_ROW + off + 1, dc), ws.cells(CHART_ROW + off + n, dc + 11)).Font.Color = RGB(60, 60, 60)

    ' 2026-09-24: split into TWO stacked charts (user request) - R1 SHORT on
    ' top with its own SELL/BUY short markers, R4 LONG below with its own
    ' SELL/BUY long markers; both carry the full 5 reference lines, their own
    ' title + legend, and their own date labels. Same hidden data block feeds
    ' both. With no R4 (not enough history) only the R1 chart is drawn.
    On Error Resume Next
    ws.ChartObjects("BIAS_CHART").Delete
    ws.ChartObjects("BIAS_CHART_R4").Delete
    On Error GoTo 0

    Dim topY As Double: topY = ws.Rows(CHART_ROW + off).Top
    Call DrawOneBiasChart(ws, "BIAS_CHART", topY, UCase(ticker) & "  R1 SHORT  (EMA" & N1 & ", RANK + TREND FILTER)", _
        "R1 SHORT", dc + 1, RGB(0, 200, 255), "SELL short", dc + 8, RGB(220, 60, 60), _
        "BUY short", dc + 9, RGB(60, 200, 90), n)
    If hasR4 Then
        Call DrawOneBiasChart(ws, "BIAS_CHART_R4", topY + BIAS_CHART_H + BIAS_CHART_GAP, _
            UCase(ticker) & "  R4 LONG  (EMA" & N4 & ", RANK + TREND FILTER)", _
            "R4 LONG", dc + 2, RR4_ACCENT, "SELL long", dc + 10, RGB(255, 120, 255), _
            "BUY long", dc + 11, RGB(120, 160, 255), n)
    End If
End Sub

Private Sub DrawOneBiasChart(ByVal ws As Worksheet, ByVal chartName As String, ByVal topY As Double, _
                              ByVal titleTxt As String, ByVal lineName As String, ByVal lineCol As Long, _
                              ByVal lineColor As Long, ByVal sellName As String, ByVal sellCol As Long, _
                              ByVal sellColor As Long, ByVal buyName As String, ByVal buyCol As Long, _
                              ByVal buyColor As Long, ByVal n As Long)
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim dc As Long: dc = CHART_DATA_COL + lc
    Dim r1 As Long: r1 = CHART_ROW + off + 1
    Dim rN As Long: rN = CHART_ROW + off + n

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(ws.Columns(CHART_COL + lc).Left, topY, BIAS_CHART_W, BIAS_CHART_H)
    co.Name = chartName
    co.Placement = xlMove
    Dim ch As Chart: Set ch = co.Chart
    ch.ChartType = xlLine
    Do While ch.SeriesCollection.count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Legend.Font.Color = CLR_MUTED: ch.Legend.Font.Size = 8
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(8, 8, 8)
    ch.PlotArea.Format.Line.Visible = msoFalse
    ch.HasTitle = True
    ch.ChartTitle.Text = titleTxt
    With ch.ChartTitle.Format.TextFrame2.TextRange.Font
        .Name = FONT_FACE: .Size = 10: .Bold = msoTrue: .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    Dim xr As Range: Set xr = ws.Range(ws.cells(r1, dc), ws.cells(rN, dc))
    Call AddBiasSeries(ch, lineName, xr, ws.Range(ws.cells(r1, lineCol), ws.cells(rN, lineCol)), lineColor, 1, False)
    Call AddBiasSeries(ch, "TOP 100", xr, ws.Range(ws.cells(r1, dc + 3), ws.cells(rN, dc + 3)), RGB(255, 255, 255), 0.75, True)
    Call AddBiasSeries(ch, "UPPER " & Format(K_THRESH * 100, "0"), xr, ws.Range(ws.cells(r1, dc + 4), ws.cells(rN, dc + 4)), RGB(220, 60, 60), 0.75, True)
    Call AddBiasSeries(ch, "ZERO", xr, ws.Range(ws.cells(r1, dc + 5), ws.cells(rN, dc + 5)), RGB(120, 120, 120), 0.75, True)
    Call AddBiasSeries(ch, "LOWER -" & Format(K_THRESH * 100, "0"), xr, ws.Range(ws.cells(r1, dc + 6), ws.cells(rN, dc + 6)), RGB(60, 160, 90), 0.75, True)
    Call AddBiasSeries(ch, "BOTTOM -100", xr, ws.Range(ws.cells(r1, dc + 7), ws.cells(rN, dc + 7)), RGB(255, 255, 255), 0.75, True)
    Call AddBiasMarkerSeries(ch, sellName, xr, ws.Range(ws.cells(r1, sellCol), ws.cells(rN, sellCol)), sellColor)
    Call AddBiasMarkerSeries(ch, buyName, xr, ws.Range(ws.cells(r1, buyCol), ws.cells(rN, buyCol)), buyColor)

    Dim ax As Axis
    Set ax = ch.Axes(xlValue)
    ax.MinimumScale = -135: ax.MaximumScale = 135
    ax.HasMajorGridlines = True
    ax.MajorGridlines.Format.Line.ForeColor.RGB = RGB(30, 30, 30)
    ax.TickLabels.Font.Color = RGB(150, 150, 150): ax.TickLabels.Font.Size = 8

    Set ax = ch.Axes(xlCategory)
    ax.TickLabels.Font.Color = RGB(150, 150, 150): ax.TickLabels.Font.Size = 7
    ax.TickLabelSpacing = CLng(Application.WorksheetFunction.Max(1, n \ 12))
End Sub

Private Sub AddBiasSeries(ByVal ch As Chart, ByVal nm As String, ByVal xr As Range, ByVal yr As Range, _
                           ByVal lineColor As Long, ByVal weight As Double, ByVal dashed As Boolean)
    Dim s As Series: Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.XValues = xr
    s.Values = yr
    s.ChartType = xlLine
    s.Format.Line.ForeColor.RGB = lineColor
    s.Format.Line.Weight = weight
    If dashed Then s.Format.Line.DashStyle = msoLineDash
    s.MarkerStyle = xlMarkerStyleNone
End Sub

' NEW (2026-09-23 second round): discrete SIGUP/SIGDN markers for the rank
' line - a line series with no line drawn, triangle markers only, plotted
' at the fixed y-offsets DrawBiasChart writes (+-112/+-124) so a marker
' never sits exactly on top of a reference line. Blank cells (every bar
' that is not itself a signal) simply plot nothing, which is what makes
' this a sparse marker series instead of a second line.
Private Sub AddBiasMarkerSeries(ByVal ch As Chart, ByVal nm As String, ByVal xr As Range, ByVal yr As Range, ByVal markerColor As Long)
    Dim s As Series: Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.XValues = xr
    s.Values = yr
    s.ChartType = xlLine
    s.Format.Line.Visible = msoFalse
    s.MarkerStyle = xlMarkerStyleTriangle
    s.MarkerSize = 7
    s.MarkerBackgroundColor = markerColor
    s.MarkerForegroundColor = markerColor
End Sub

' ----------------------------------------------------------------
'  K-line + volume (2026-09-24). Two pairs, right of the R charts:
'  R1 SHORT pair carries EMA(N1) and the short SELL/BUY triangles, R4
'  LONG pair carries EMA(N4) and the long ones. Candles are a plain line
'  chart group with hi-lo lines + up/down bars (exactly what Excel's
'  native OHLC stock chart is underneath); EMA and the signal triangles
'  live in a SECONDARY group pinned to the same fixed axis scale, because
'  hi-lo lines span every series in their own group and up/down bars use
'  the group's first and last series - extra series in that group would
'  corrupt both. Red up / green down (TW convention, as on the RR4 page).
' ----------------------------------------------------------------

' JSON numeric array between "<key>":[ and the next ] (null -> "null").
Private Function BiasJsonArr(ByRef resp As String, ByVal key As String, ByVal fromPos As Long, ByRef parts() As String) As Boolean
    Dim p1 As Long, p2 As Long
    p1 = InStr(fromPos, resp, """" & key & """:[")
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(key) + 4
    p2 = InStr(p1, resp, "]")
    If p2 = 0 Then Exit Function
    parts = Split(Mid(resp, p1, p2 - p1), ",")
    BiasJsonArr = True
End Function

' One Yahoo symbol, 10y daily OHLCV. O/H/L are scaled by adjclose/close (what
' yfinance auto_adjust does) so they sit on the same adjusted basis as the
' closes the EMAs and R lines are computed from; C is adjclose; V is raw.
' Rows with any null field are dropped. Returns the row count (0 = failed).
Private Function FetchBiasOhlcvRaw(ByVal sym As String, ByRef days() As Long, ByRef o() As Double, _
                                    ByRef h() As Double, ByRef l() As Double, ByRef c() As Double, _
                                    ByRef v() As Double) As Long
    Dim http As Object, resp As String, status As Long
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://query1.finance.yahoo.com/v8/finance/chart/" & sym & "?range=10y&interval=1d", False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    status = http.Status
    resp = http.responseText
    On Error GoTo 0
    If status <> 200 Or Len(resp) = 0 Then Exit Function

    Dim ts() As String, op() As String, hi() As String, lo() As String, cl() As String, vo() As String, ac() As String
    If Not BiasJsonArr(resp, "timestamp", 1, ts) Then Exit Function
    If Not BiasJsonArr(resp, "open", 1, op) Then Exit Function
    If Not BiasJsonArr(resp, "high", 1, hi) Then Exit Function
    If Not BiasJsonArr(resp, "low", 1, lo) Then Exit Function
    If Not BiasJsonArr(resp, "close", 1, cl) Then Exit Function       ' "adjclose":[ does not match "close":[
    If Not BiasJsonArr(resp, "volume", 1, vo) Then Exit Function
    Dim pAdj As Long: pAdj = InStr(resp, """adjclose"":[")            ' the first one opens the object
    If pAdj = 0 Then Exit Function
    If Not BiasJsonArr(resp, "adjclose", pAdj + 1, ac) Then Exit Function

    Dim n As Long: n = UBound(ts) + 1
    If UBound(op) + 1 < n Then n = UBound(op) + 1
    If UBound(hi) + 1 < n Then n = UBound(hi) + 1
    If UBound(lo) + 1 < n Then n = UBound(lo) + 1
    If UBound(cl) + 1 < n Then n = UBound(cl) + 1
    If UBound(vo) + 1 < n Then n = UBound(vo) + 1
    If UBound(ac) + 1 < n Then n = UBound(ac) + 1
    If n <= 0 Then Exit Function
    ReDim days(0 To n - 1): ReDim o(0 To n - 1): ReDim h(0 To n - 1)
    ReDim l(0 To n - 1): ReDim c(0 To n - 1): ReDim v(0 To n - 1)

    Dim k As Long, cnt As Long, rawC As Double, ratio As Double
    For k = 0 To n - 1
        If IsNumeric(Trim(ts(k))) And IsNumeric(Trim(op(k))) And IsNumeric(Trim(hi(k))) And _
           IsNumeric(Trim(lo(k))) And IsNumeric(Trim(cl(k))) And IsNumeric(Trim(vo(k))) And _
           IsNumeric(Trim(ac(k))) Then
            rawC = CDbl(Val(cl(k)))
            If rawC > 0 Then
                ratio = CDbl(Val(ac(k))) / rawC
                days(cnt) = CLng(Val(ts(k)) \ 86400)
                o(cnt) = CDbl(Val(op(k))) * ratio
                h(cnt) = CDbl(Val(hi(k))) * ratio
                l(cnt) = CDbl(Val(lo(k))) * ratio
                c(cnt) = CDbl(Val(ac(k)))
                v(cnt) = CDbl(Val(vo(k)))
                cnt = cnt + 1
            End If
        End If
    Next k
    FetchBiasOhlcvRaw = cnt
End Function

' OHLCV aligned 1:1 to the chart dates (0-based, length n, newest = last).
' A bare TW number is tried as .TW then .TWO (same rule as
' modvolatility.GetHistoricalData). A chart bar with no OHLC row on that date
' falls back to a flat bar at the adjusted close with volume 0.
Private Function FetchBiasOhlc(ByVal ticker As String, ByRef chartDates() As Date, ByVal n As Long, _
                                ByRef closeArr() As Double, ByVal cnt As Long, _
                                ByRef oo() As Double, ByRef hh() As Double, ByRef ll() As Double, _
                                ByRef cc() As Double, ByRef vv() As Double) As Boolean
    Dim t As String: t = UCase$(Trim$(ticker))
    Dim d() As Long, ro() As Double, rh() As Double, rl() As Double, rc() As Double, rv() As Double
    Dim m As Long
    If InStr(t, ".") = 0 And IsNumeric(t) Then
        m = FetchBiasOhlcvRaw(t & ".TW", d, ro, rh, rl, rc, rv)
        If m = 0 Then m = FetchBiasOhlcvRaw(t & ".TWO", d, ro, rh, rl, rc, rv)
    Else
        m = FetchBiasOhlcvRaw(t, d, ro, rh, rl, rc, rv)
    End If
    If m = 0 Then Exit Function

    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To m - 1
        idx(d(i)) = i
    Next i

    ReDim oo(0 To n - 1): ReDim hh(0 To n - 1): ReDim ll(0 To n - 1): ReDim cc(0 To n - 1): ReDim vv(0 To n - 1)
    Dim dayKey As Long, j As Long, px As Double
    For i = 0 To n - 1
        dayKey = CLng(Int(CDbl(chartDates(i)))) - 25569
        px = closeArr(cnt - n + i)
        If idx.Exists(dayKey) Then
            j = idx(dayKey)
            oo(i) = ro(j): hh(i) = rh(j): ll(i) = rl(j): cc(i) = px: vv(i) = rv(j)
            If hh(i) < px Then hh(i) = px
            If ll(i) > px Then ll(i) = px
        Else
            oo(i) = px: hh(i) = px: ll(i) = px: cc(i) = px: vv(i) = 0
        End If
    Next i
    FetchBiasOhlc = True
End Function

' 1 / 2 / 2.5 / 5 x 10^k step that gives at most 8 gridlines over span.
Private Function NiceStep(ByVal span As Double) As Double
    If span <= 0 Then NiceStep = 1: Exit Function
    Dim mag As Double: mag = 10 ^ Int(Log(span / 6) / Log(10))
    Dim mult As Variant: mult = Array(1, 2, 2.5, 5, 10)
    Dim k As Long
    For k = 0 To 4
        If span / (mag * mult(k)) <= 8 Then NiceStep = mag * mult(k): Exit Function
    Next k
    NiceStep = mag * 10
End Function

Private Sub DrawBiasKlines(ByVal ws As Worksheet, ByVal ticker As String, ByVal n As Long, _
                            ByRef ko() As Double, ByRef kh() As Double, ByRef kl() As Double, _
                            ByRef kc() As Double, ByRef kv() As Double, _
                            ByRef emaS() As Double, ByRef emaL() As Double, ByVal hasR4 As Boolean, _
                            ByRef sigUp1() As Boolean, ByRef sigDn1() As Boolean, _
                            ByRef sigUp4() As Boolean, ByRef sigDn4() As Boolean)
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim dc As Long: dc = CHART_DATA_COL + lc
    Dim kb As Long: kb = dc + KB_OFF
    Dim r0 As Long: r0 = CHART_ROW + off + 1

    Dim hdr As Variant
    hdr = Array("O", "H", "L", "C", "EMA" & N1, "EMA" & N4, "SELLS", "BUYS", "SELLL", "BUYL", "VOLUP", "VOLDN")
    Dim j As Long
    For j = 0 To 11
        ws.cells(CHART_ROW + off, kb + j).Value = hdr(j)
    Next j
    ws.Range(ws.cells(CHART_ROW + off, kb), ws.cells(CHART_ROW + off, kb + 11)).Font.Color = RGB(90, 90, 90)

    ' one array write for the whole block; blank = not plotted (signal / off-side volume)
    Dim blk() As Variant: ReDim blk(1 To n, 1 To 12)
    Dim i As Long, r As Long
    Dim lo As Double, hi As Double: lo = kl(0): hi = kh(0)
    For i = 0 To n - 1
        r = i + 1
        blk(r, 1) = ko(i): blk(r, 2) = kh(i): blk(r, 3) = kl(i): blk(r, 4) = kc(i)
        blk(r, 5) = emaS(i)
        If hasR4 Then blk(r, 6) = emaL(i)
        If sigUp1(i) Then blk(r, 7) = kh(i) * 1.025
        If sigDn1(i) Then blk(r, 8) = kl(i) * 0.975
        If hasR4 Then
            If sigUp4(i) Then blk(r, 9) = kh(i) * 1.025
            If sigDn4(i) Then blk(r, 10) = kl(i) * 0.975
        End If
        If kc(i) >= ko(i) Then blk(r, 11) = kv(i) Else blk(r, 12) = kv(i)

        If kl(i) < lo Then lo = kl(i)
        If kh(i) > hi Then hi = kh(i)
        If emaS(i) < lo Then lo = emaS(i)
        If emaS(i) > hi Then hi = emaS(i)
        If hasR4 Then
            If emaL(i) < lo Then lo = emaL(i)
            If emaL(i) > hi Then hi = emaL(i)
        End If
    Next i
    With ws.Range(ws.cells(r0, kb), ws.cells(r0 + n - 1, kb + 11))
        .Value = blk
        .Font.Color = RGB(60, 60, 60)
    End With

    ' shared fixed price scale (primary and the hidden secondary axis must match)
    lo = lo * 0.96: hi = hi * 1.04            ' room for the +-2.5% signal triangles
    Dim stp As Double: stp = NiceStep(hi - lo)
    Dim axMin As Double, axMax As Double
    axMin = Int(lo / stp) * stp
    axMax = -Int(-hi / stp) * stp
    Dim axFmt As String: axFmt = IIf(stp < 1, "#,##0.00", "#,##0")

    Dim leftX As Double: leftX = ws.Columns(CHART_COL + lc).Left + BIAS_CHART_W + KLINE_GAP
    Dim topY As Double: topY = ws.Rows(CHART_ROW + off).Top
    Dim tk As String: tk = UCase(ticker)

    Call DrawOneKline(ws, "BIAS_K_R1", leftX, topY, tk & "  K-LINE  (SHORT: EMA" & N1 & ")", dc, kb, 4, _
        "EMA" & N1, RGB(0, 200, 255), 6, "SELL short", RGB(220, 60, 60), 7, "BUY short", RGB(60, 200, 90), _
        n, axMin, axMax, stp, axFmt)
    Call DrawOneVolume(ws, "BIAS_V_R1", leftX, topY + KLINE_H + KV_GAP, dc, kb, n)
    If hasR4 Then
        Dim topY4 As Double: topY4 = topY + BIAS_CHART_H + BIAS_CHART_GAP
        Call DrawOneKline(ws, "BIAS_K_R4", leftX, topY4, tk & "  K-LINE  (LONG: EMA" & N4 & ")", dc, kb, 5, _
            "EMA" & N4, RR4_ACCENT, 8, "SELL long", RGB(255, 120, 255), 9, "BUY long", RGB(120, 160, 255), _
            n, axMin, axMax, stp, axFmt)
        Call DrawOneVolume(ws, "BIAS_V_R4", leftX, topY4 + KLINE_H + KV_GAP, dc, kb, n)
    End If
End Sub

' kb + emaIdx / sellIdx / buyIdx are column offsets inside the K data block.
Private Sub DrawOneKline(ByVal ws As Worksheet, ByVal chartName As String, ByVal leftX As Double, _
                          ByVal topY As Double, ByVal titleTxt As String, ByVal dc As Long, ByVal kb As Long, _
                          ByVal emaIdx As Long, ByVal emaName As String, ByVal emaColor As Long, _
                          ByVal sellIdx As Long, ByVal sellName As String, ByVal sellColor As Long, _
                          ByVal buyIdx As Long, ByVal buyName As String, ByVal buyColor As Long, _
                          ByVal n As Long, ByVal axMin As Double, ByVal axMax As Double, _
                          ByVal stp As Double, ByVal axFmt As String)
    Dim off As Long: off = NavOffset(ws)
    Dim r1 As Long: r1 = CHART_ROW + off + 1
    Dim rN As Long: rN = CHART_ROW + off + n
    Dim upClr As Long: upClr = RGB(220, 60, 60)
    Dim dnClr As Long: dnClr = RGB(255, 255, 255)   ' 2026-09-25: red up / white down, black borders

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(leftX, topY, BIAS_CHART_W, KLINE_H)
    co.Name = chartName
    co.Placement = xlMove
    Dim ch As Chart: Set ch = co.Chart
    ch.ChartType = xlLine
    Do While ch.SeriesCollection.count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Legend.Font.Color = CLR_MUTED: ch.Legend.Font.Size = 8
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(8, 8, 8)
    ch.PlotArea.Format.Line.Visible = msoFalse
    ch.HasTitle = True
    ch.ChartTitle.Text = titleTxt
    With ch.ChartTitle.Format.TextFrame2.TextRange.Font
        .Name = FONT_FACE: .Size = 10: .Bold = msoTrue: .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    Dim xr As Range: Set xr = ws.Range(ws.cells(r1, dc), ws.cells(rN, dc))
    Dim nm As Variant, k As Long
    nm = Array("OPEN", "HIGH", "LOW", "CLOSE")
    For k = 0 To 3
        Call AddBiasSeries(ch, CStr(nm(k)), xr, ws.Range(ws.cells(r1, kb + k), ws.cells(rN, kb + k)), RGB(255, 255, 255), 1, False)
        ch.SeriesCollection(ch.SeriesCollection.count).Format.Line.Visible = msoFalse
    Next k
    With ch.ChartGroups(1)
        .HasHiLoLines = True
        .HasUpDownBars = True
        .GapWidth = 10                      ' wide bodies: the 0.75pt black borders eat thin ones
        .HiLoLines.Format.Line.ForeColor.RGB = RGB(255, 255, 255)
        .HiLoLines.Format.Line.Weight = 0.75
        .UpBars.Format.Fill.ForeColor.RGB = upClr
        .UpBars.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .UpBars.Format.Line.Weight = 0.75
        .DownBars.Format.Fill.ForeColor.RGB = dnClr
        .DownBars.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .DownBars.Format.Line.Weight = 0.75
    End With

    ' EMA + signal triangles: secondary group, same fixed scale as the primary axis
    Call AddBiasSeries(ch, emaName, xr, ws.Range(ws.cells(r1, kb + emaIdx), ws.cells(rN, kb + emaIdx)), emaColor, 1.25, False)
    ch.SeriesCollection(ch.SeriesCollection.count).AxisGroup = xlSecondary
    With ch.SeriesCollection(ch.SeriesCollection.count).Format.Line      ' moving groups resets the line colour to black
        .Visible = msoTrue: .ForeColor.RGB = emaColor: .Weight = 1.25
    End With
    Call AddBiasMarkerSeries(ch, sellName, xr, ws.Range(ws.cells(r1, kb + sellIdx), ws.cells(rN, kb + sellIdx)), sellColor)
    ch.SeriesCollection(ch.SeriesCollection.count).AxisGroup = xlSecondary
    Call AddBiasMarkerSeries(ch, buyName, xr, ws.Range(ws.cells(r1, kb + buyIdx), ws.cells(rN, kb + buyIdx)), buyColor)
    ch.SeriesCollection(ch.SeriesCollection.count).AxisGroup = xlSecondary

    Dim ax As Axis
    Set ax = ch.Axes(xlValue, xlPrimary)
    ax.MinimumScale = axMin: ax.MaximumScale = axMax: ax.MajorUnit = stp
    ax.HasMajorGridlines = True
    ax.MajorGridlines.Format.Line.ForeColor.RGB = RGB(30, 30, 30)
    ax.TickLabels.NumberFormat = axFmt
    ax.TickLabels.Font.Color = RGB(150, 150, 150): ax.TickLabels.Font.Size = 8

    Set ax = ch.Axes(xlValue, xlSecondary)
    ax.MinimumScale = axMin: ax.MaximumScale = axMax: ax.MajorUnit = stp
    ax.HasMajorGridlines = False
    ax.TickLabelPosition = xlTickLabelPositionNone
    ax.MajorTickMark = xlTickMarkNone
    ax.Format.Line.Visible = msoFalse

    Set ax = ch.Axes(xlCategory, xlPrimary)
    ax.CategoryType = xlCategoryScale                    ' bar index, no weekend gaps
    ax.TickLabelPosition = xlTickLabelPositionNone       ' dates are shown on the volume chart below
    ax.MajorTickMark = xlTickMarkNone

    ' the four candle series are structural, not something to list in the legend
    On Error Resume Next
    For k = 1 To 4
        ch.Legend.LegendEntries(1).Delete
    Next k
    ch.PlotArea.InsideWidth = BIAS_CHART_W - K_PLOT_LEFT - K_PLOT_RIGHT
    ch.PlotArea.InsideLeft = K_PLOT_LEFT              ' width first: setting it after can shift the left edge
    ch.PlotArea.InsideTop = 28
    ch.PlotArea.InsideHeight = KLINE_H - 28 - 30
    On Error GoTo 0
End Sub

Private Sub DrawOneVolume(ByVal ws As Worksheet, ByVal chartName As String, ByVal leftX As Double, _
                           ByVal topY As Double, ByVal dc As Long, ByVal kb As Long, ByVal n As Long)
    Dim off As Long: off = NavOffset(ws)
    Dim r1 As Long: r1 = CHART_ROW + off + 1
    Dim rN As Long: rN = CHART_ROW + off + n

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(leftX, topY, BIAS_CHART_W, VOL_H)
    co.Name = chartName
    co.Placement = xlMove
    Dim ch As Chart: Set ch = co.Chart
    ch.ChartType = xlColumnClustered
    Do While ch.SeriesCollection.count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.HasLegend = False
    ch.HasTitle = False
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(8, 8, 8)
    ch.PlotArea.Format.Line.Visible = msoFalse

    Dim xr As Range: Set xr = ws.Range(ws.cells(r1, dc), ws.cells(rN, dc))
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "VOL UP"
    s.XValues = xr
    s.Values = ws.Range(ws.cells(r1, kb + 10), ws.cells(rN, kb + 10))
    s.Format.Fill.ForeColor.RGB = RGB(220, 60, 60)
    s.Format.Line.Visible = msoTrue: s.Format.Line.ForeColor.RGB = RGB(0, 0, 0): s.Format.Line.Weight = 0.75
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "VOL DOWN"
    s.XValues = xr
    s.Values = ws.Range(ws.cells(r1, kb + 11), ws.cells(rN, kb + 11))
    s.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.Format.Line.Visible = msoTrue: s.Format.Line.ForeColor.RGB = RGB(0, 0, 0): s.Format.Line.Weight = 0.75
    ch.ChartGroups(1).Overlap = 100
    ch.ChartGroups(1).GapWidth = 40

    Dim ax As Axis
    Set ax = ch.Axes(xlValue)
    ax.HasMajorGridlines = True
    ax.MajorGridlines.Format.Line.ForeColor.RGB = RGB(30, 30, 30)
    ax.TickLabels.NumberFormat = "[>=1000000]0.0,,""M"";[>=1000]0,""K"";0"
    ax.TickLabels.Font.Color = RGB(150, 150, 150): ax.TickLabels.Font.Size = 8

    Set ax = ch.Axes(xlCategory)
    ax.CategoryType = xlCategoryScale
    ax.TickLabels.Font.Color = RGB(150, 150, 150): ax.TickLabels.Font.Size = 7
    ax.TickLabelSpacing = CLng(Application.WorksheetFunction.Max(1, n \ 12))

    On Error Resume Next
    ch.PlotArea.InsideWidth = BIAS_CHART_W - K_PLOT_LEFT - K_PLOT_RIGHT
    ch.PlotArea.InsideLeft = K_PLOT_LEFT              ' width first: setting it after can shift the left edge
    ch.PlotArea.InsideTop = 6
    ch.PlotArea.InsideHeight = VOL_H - 6 - 24
    On Error GoTo 0
End Sub

' Explicitly re-blacks the hidden data block after clearing it - a bare
' Range.Clear resets cell formatting to Excel's default (white/no fill),
' which made the block flash white on every rebuild (2026-09-24 fix).
Private Sub ClearBiasChart(ByVal ws As Worksheet)
    On Error Resume Next
    ws.ChartObjects("BIAS_CHART").Delete
    ws.ChartObjects("BIAS_CHART_R4").Delete
    ws.ChartObjects("BIAS_K_R1").Delete
    ws.ChartObjects("BIAS_V_R1").Delete
    ws.ChartObjects("BIAS_K_R4").Delete
    ws.ChartObjects("BIAS_V_R4").Delete
    On Error GoTo 0
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim dc As Long: dc = CHART_DATA_COL + lc
    Dim rng As Range
    Set rng = ws.Range(ws.cells(CHART_ROW + off, dc), ws.cells(CHART_ROW + off + NEED_BARS + 5, dc + KB_LAST))
    rng.ClearContents
    rng.Interior.Color = RGB(0, 0, 0)
    ' the block moved to ZA on 2026-09-25: wipe the old AE-based copy (same shape, same rows)
    Dim ldc As Long: ldc = LEGACY_DATA_COL + lc
    Set rng = ws.Range(ws.cells(CHART_ROW + off, ldc), ws.cells(CHART_ROW + off + NEED_BARS + 5, ldc + KB_LAST))
    rng.ClearContents
    rng.Interior.Color = RGB(0, 0, 0)
End Sub

' ----------------------------------------------------------------
'  Sheet / shell
' ----------------------------------------------------------------
Private Function EnsureBiasSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(BIAS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Dim prev As Object: Set prev = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = BIAS_SHEET
        ActiveWindow.DisplayGridlines = False
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If
    Set EnsureBiasSheet = ws
End Function

Private Sub DrawShell(ByVal ws As Worksheet, ByVal keepTk As String)
    ws.cells.Clear
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = FONT_FACE
        .Font.Size = 9
        .Font.Color = CLR_TEXT
        .VerticalAlignment = xlCenter
        .RowHeight = 16
    End With
    With ws.cells(PG_TITLE, 1)
        .Value = "BIAS"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 14
    End With
    ws.Rows(PG_TITLE).RowHeight = 24

    Call StackLabel(ws.cells(PG_LBL, COL_TK), "TICKER <GO>")
    ws.Rows(PG_LBL).RowHeight = 18
    ws.Rows(PG_IN).RowHeight = 20
    Call InputCell(ws.cells(PG_IN, COL_TK), keepTk)

    ws.Columns(1).ColumnWidth = 14
    Dim c As Long
    For c = 2 To 8: ws.Columns(c).ColumnWidth = 12: Next c
End Sub

Private Sub StackLabel(ByVal cell As Range, ByVal txt As String)
    With cell
        .Value = txt
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
End Sub

Private Sub InputCell(ByVal cell As Range, ByVal v As String)
    With cell
        .NumberFormat = "@"
        .Value = v
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub FinishPage(ByVal ws As Worksheet)
    On Error Resume Next
    Call NavAdd(ws, "B")
    Call EnsureSheetCode(ws)
    On Error GoTo 0
End Sub

' Sheet event code (mirrors modEarnings.EnsureSheetCode), written into
' the document module the first time the page is built. Needs "Trust
' access to the VBA project object model", like the RRG / Earnings /
' Thesis pages.
Private Sub EnsureSheetCode(ByVal ws As Worksheet)
    On Error GoTo Skip
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 Then
            If comp.Properties("Name").Value = ws.Name Then
                Dim cm As Object: Set cm = comp.CodeModule
                If cm.CountOfLines > 0 Then
                    If InStr(cm.Lines(1, cm.CountOfLines), "BiasChange") > 0 Then Exit Sub
                    cm.DeleteLines 1, cm.CountOfLines
                End If
                cm.AddFromString "Option Explicit" & vbCrLf & vbCrLf & _
                    "' Bias page: TICKER input -> modBias.BiasChange (see modBias.bas)" & vbCrLf & _
                    "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
                    "    Call BiasChange(Me, Target)" & vbCrLf & _
                    "End Sub" & vbCrLf
                Exit Sub
            End If
        End If
    Next comp
    Exit Sub
Skip:
    Call NavNotify("Bias page built, but its sheet event code could not be written (enable Trust access to the VBA project object model, or paste the Worksheet_Change block by hand)", True)
End Sub

' ----------------------------------------------------------------
'  Small local helpers (deliberately duplicated rather than shared -
'  same call this codebase already makes for modLivePrice etc.)
' ----------------------------------------------------------------
Private Function CellStr(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    CellStr = Trim(CStr(v))
End Function
