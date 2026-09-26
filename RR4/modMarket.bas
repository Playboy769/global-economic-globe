Attribute VB_Name = "modMarket"
Option Explicit

' ============================================================================
' modMarket -- market-condition snapshot (data layer only; the RR4 page draws it)
'
'   Public Function MarketSnapshot() As Variant   -> m(1 To 7, 1 To 9)
'   Public Function MarketLabel(i As Long) As String
'
' Rows (fixed order):
'   1 FEAR & GREED       CNN Fear & Greed Index, 0-100
'   2 VIX                Yahoo ^VIX
'   3 SOX                Yahoo ^SOX (index level)
'   4 SOX VOL 20D        20-day realized vol of SOX, annualised, in %
'   5 SOX VOL 20D AVG    mean of the last 20 daily values of row 4, in %
'   6 BRENT SPOT         EIA spot via FRED DCOILBRENTEU (USD/bbl)
'   7 WTI SPOT           EIA spot via FRED DCOILWTICO   (USD/bbl)
'
' Columns (per row):
'   1 last        latest value
'   2 chgDay      latest minus previous observation, in the series' own unit
'                 (points for F&G / VIX / SOX / vols, USD for oil)
'   3 chgDayPct   (last / previous - 1) * 100, relative %, for every row
'   4 chg1W       vs 5 observations back  (unit says which: see col 9)
'   5 chg1M       vs 21 observations back (same unit rule)
'   6 pctile      percent of the last 252 observations (incl. today; fewer if
'                 the history is shorter) that are <= the latest value, 0-100
'   7 asOf        date of the last observation, "yyyy-mm-dd" (oil lags 1-2 days)
'   8 note        F&G: CNN rating text (Fear / Extreme Greed ...); others ""
'   9 unit        unit of columns 4 and 5:  "pt" = point difference
'                 (F&G, SOX VOL 20D, SOX VOL 20D AVG), "%" = percent change
'                 (VIX, SOX, BRENT, WTI)
'
' Definitions:
'   * "observation" = a valid data point (FRED holidays "." are skipped, Yahoo
'     null closes are skipped). 1W = 5 observations back, 1M = 21 back.
'   * Realized vol: sample stdev (n-1) of the last 20 daily LOG returns of the
'     SOX adjusted close * sqrt(252) * 100. The series of daily vols is built
'     over the whole 2y history; row 5 is the plain mean of its last 20 values
'     and its own series (rolling 20-day mean) drives the percentile / 1W / 1M.
'   * F&G "today" = fear_and_greed.score; history = graphdata history,
'     de-duplicated to one value per UTC day (last entry of a day wins; if the
'     last history day equals the score's date it is replaced by the score).
'
' Failure policy: every indicator is isolated. Anything that cannot be fetched
' or parsed leaves its whole row Empty (the caller prints "-"). Nothing is
' cached and nothing is written to any worksheet. HTTP timeouts are short and
' a request is retried at most once (network error / 429 / 5xx only).
'
' Source quirks (measured 2026-09-27):
'   * CNN answers 418 unless User-Agent is a full browser string AND a
'     Referer of edition.cnn.com is sent.
'   * FRED stalls (no answer) when the User-Agent looks like a browser, so
'     FRED gets a plain non-browser UA; cosd= limits the CSV to ~500 days.
'   * Yahoo chart accepts any UA; '^' must be encoded as %5E.
' Self-contained: depends on no other RR4 module. Keep this file pure ASCII.
' ============================================================================

Private Const MK_ROWS As Long = 7
Private Const MK_COLS As Long = 9

Private Const UA_BROWSER As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
Private Const UA_PLAIN As String = "curl/8.4.0"

Private Const T_RESOLVE_MS As Long = 5000
Private Const T_CONNECT_MS As Long = 5000
Private Const T_SEND_MS As Long = 8000
Private Const T_RECEIVE_MS As Long = 10000

Private Const RETRY_BUDGET_SEC As Double = 40#

Private mT0 As Double

Public Function MarketLabel(ByVal i As Long) As String
    Select Case i
        Case 1: MarketLabel = "FEAR & GREED"
        Case 2: MarketLabel = "VIX"
        Case 3: MarketLabel = "SOX"
        Case 4: MarketLabel = "SOX VOL 20D"
        Case 5: MarketLabel = "SOX VOL 20D AVG"
        Case 6: MarketLabel = "BRENT SPOT"
        Case 7: MarketLabel = "WTI SPOT"
        Case Else: MarketLabel = ""
    End Select
End Function

Public Function MarketSnapshot() As Variant
    Dim m(1 To MK_ROWS, 1 To MK_COLS) As Variant
    mT0 = Timer

    On Error Resume Next
    LoadFearGreed m
    Err.Clear
    LoadVix m
    Err.Clear
    LoadSoxFamily m
    Err.Clear
    LoadOil m, 6, "DCOILBRENTEU"
    Err.Clear
    LoadOil m, 7, "DCOILWTICO"
    Err.Clear
    On Error GoTo 0

    MarketSnapshot = m
End Function

' ---------------------------------------------------------------- indicators

Private Sub LoadFearGreed(ByRef m() As Variant)
    On Error GoTo Fail
    Dim url As String, s As String
    url = "https://production.dataviz.cnn.io/index/fearandgreed/graphdata/" & DateStr(Date - 400)
    s = HttpGetText(url, UA_BROWSER, "https://edition.cnn.com/", "https://edition.cnn.com")
    If Len(s) = 0 Then Exit Sub

    ' --- today's block: "fear_and_greed": {"score": ..., "rating": ..., "timestamp": ...}
    Dim p As Long, pHist As Long
    p = InStr(1, s, """fear_and_greed""", vbBinaryCompare)
    pHist = InStr(1, s, """fear_and_greed_historical""", vbBinaryCompare)
    If p = 0 Or pHist = 0 Then Exit Sub

    Dim pS As Long, score As Double, rating As String, tsStr As String
    pS = InStr(p, s, """score""", vbBinaryCompare)
    If pS = 0 Or pS > pHist Then Exit Sub
    score = NumAfterKey(s, pS + 7)
    rating = StrAfterKey(s, "rating", p, pHist)
    tsStr = StrAfterKey(s, "timestamp", p, pHist)

    ' --- history entries: {"x": ms, "y": value, "rating": ...}
    Dim pData As Long
    pData = InStr(pHist, s, """data""", vbBinaryCompare)
    If pData = 0 Then Exit Sub
    Dim pEnd As Long
    pEnd = InStr(pData, s, "]", vbBinaryCompare)      ' entries hold no nested arrays
    If pEnd = 0 Then Exit Sub

    Dim v() As Double, d() As Double, n As Long
    ReDim v(1 To 800)
    ReDim d(1 To 800)
    Dim pos As Long, px As Long, py As Long, xv As Double, yv As Double, dayNo As Double
    pos = pData
    Do
        px = InStr(pos, s, """x""", vbBinaryCompare)
        If px = 0 Or px > pEnd Then Exit Do
        py = InStr(px, s, """y""", vbBinaryCompare)
        If py = 0 Or py > pEnd Then Exit Do
        xv = NumAfterKey(s, px + 3)
        yv = NumAfterKey(s, py + 3)
        dayNo = Fix(xv / 86400000#)
        If n > 0 Then
            If d(n) = dayNo Then
                v(n) = yv                       ' same UTC day: last entry wins
                pos = py + 3
                GoTo NextEntry
            End If
        End If
        If n >= UBound(v) Then Exit Do
        n = n + 1
        d(n) = dayNo
        v(n) = yv
        pos = py + 3
NextEntry:
    Loop

    ' --- splice today's score
    If Len(tsStr) >= 10 Then
        Dim todayDay As Double
        todayDay = DateValueIso(Left$(tsStr, 10))
        If todayDay > 0 Then
            todayDay = todayDay - EpochDay()
            If n > 0 Then
                If d(n) = todayDay Then
                    v(n) = score
                ElseIf d(n) < todayDay And n < UBound(v) Then
                    n = n + 1
                    d(n) = todayDay
                    v(n) = score
                End If
            End If
        End If
    End If
    If n < 2 Then Exit Sub

    FillStats m, 1, v, n, "pt", DayToStr(d(n)), TitleCase(rating)
    Exit Sub
Fail:
    ClearRow m, 1
End Sub

Private Sub LoadVix(ByRef m() As Variant)
    On Error GoTo Fail
    Dim d() As Double, v() As Double, n As Long
    If Not FetchYahoo("%5EVIX", d, v, n) Then Exit Sub
    FillStats m, 2, v, n, "%", DayToStr(d(n)), ""
    Exit Sub
Fail:
    ClearRow m, 2
End Sub

Private Sub LoadSoxFamily(ByRef m() As Variant)
    ' SOX is fetched ONCE and feeds rows 3, 4 and 5.
    On Error GoTo Fail
    Dim d() As Double, v() As Double, n As Long
    If Not FetchYahoo("%5ESOX", d, v, n) Then Exit Sub

    On Error Resume Next
    FillStats m, 3, v, n, "%", DayToStr(d(n)), ""
    If Err.Number <> 0 Then ClearRow m, 3
    Err.Clear
    On Error GoTo Fail

    ' daily 20d realized vol series (annualised, %)
    Dim nr As Long, i As Long, j As Long
    If n < 22 Then Exit Sub
    Dim lr() As Double
    ReDim lr(1 To n - 1)
    For i = 2 To n
        If v(i) > 0 And v(i - 1) > 0 Then
            lr(i - 1) = Log(v(i) / v(i - 1))
        Else
            Exit Sub
        End If
    Next i
    nr = n - 1

    Dim nv As Long
    nv = nr - 19
    If nv < 22 Then Exit Sub
    Dim vol() As Double
    ReDim vol(1 To nv)
    Dim sm As Double, mean As Double, ss As Double
    For i = 1 To nv
        sm = 0#
        For j = i To i + 19
            sm = sm + lr(j)
        Next j
        mean = sm / 20#
        ss = 0#
        For j = i To i + 19
            ss = ss + (lr(j) - mean) * (lr(j) - mean)
        Next j
        vol(i) = Sqr(ss / 19#) * Sqr(252#) * 100#
    Next i

    ' vol(i) belongs to price index i + 20 (returns i..i+19 end at price i+20)
    FillStats m, 4, vol, nv, "pt", DayToStr(d(n)), ""

    Dim na As Long
    na = nv - 19
    If na >= 22 Then
        Dim av() As Double
        ReDim av(1 To na)
        For i = 1 To na
            sm = 0#
            For j = i To i + 19
                sm = sm + vol(j)
            Next j
            av(i) = sm / 20#
        Next i
        FillStats m, 5, av, na, "pt", DayToStr(d(n)), ""
    End If
    Exit Sub
Fail:
    ClearRow m, 4
    ClearRow m, 5
End Sub

Private Sub LoadOil(ByRef m() As Variant, ByVal r As Long, ByVal seriesId As String)
    On Error GoTo Fail
    Dim s As String
    s = HttpGetText("https://fred.stlouisfed.org/graph/fredgraph.csv?id=" & seriesId & "&cosd=" & DateStr(Date - 500), UA_PLAIN, "", "")
    If Len(s) = 0 Then Exit Sub

    Dim lines() As String
    s = Replace(s, vbCr, "")
    lines = Split(s, vbLf)

    Dim v() As Double, d() As Double, n As Long, i As Long
    ReDim v(1 To UBound(lines) + 1)
    ReDim d(1 To UBound(lines) + 1)
    Dim parts() As String, dt As Double
    For i = 1 To UBound(lines)                  ' line 0 is the header
        If Len(lines(i)) > 8 Then
            parts = Split(lines(i), ",")
            If UBound(parts) >= 1 Then
                If Len(parts(1)) > 0 And parts(1) <> "." Then
                    dt = DateValueIso(Left$(parts(0), 10))
                    If dt > 0 Then
                        n = n + 1
                        d(n) = dt - EpochDay()
                        v(n) = Val(parts(1))
                    End If
                End If
            End If
        End If
    Next i
    If n < 2 Then Exit Sub

    FillStats m, r, v, n, "%", DayToStr(d(n)), ""
    Exit Sub
Fail:
    ClearRow m, r
End Sub

' ---------------------------------------------------------------- statistics

' v(1..n) chronological. mode "pt": 1W/1M as point difference, "%": as percent.
Private Sub FillStats(ByRef m() As Variant, ByVal r As Long, ByRef v() As Double, _
                      ByVal n As Long, ByVal mode As String, ByVal asOf As String, _
                      ByVal note As String)
    If n < 2 Then Exit Sub
    Dim last As Double, prev As Double
    last = v(n)
    prev = v(n - 1)

    m(r, 1) = last
    m(r, 2) = last - prev
    If prev <> 0 Then m(r, 3) = (last / prev - 1#) * 100#
    If n >= 6 Then m(r, 4) = Change(last, v(n - 5), mode)
    If n >= 22 Then m(r, 5) = Change(last, v(n - 21), mode)

    Dim lo As Long, i As Long, cnt As Long, tot As Long
    lo = n - 251
    If lo < 1 Then lo = 1
    For i = lo To n
        tot = tot + 1
        If v(i) <= last Then cnt = cnt + 1
    Next i
    If tot > 0 Then m(r, 6) = cnt / tot * 100#

    m(r, 7) = asOf
    m(r, 8) = note
    m(r, 9) = mode
End Sub

Private Function Change(ByVal last As Double, ByVal base As Double, ByVal mode As String) As Variant
    If mode = "pt" Then
        Change = last - base
    ElseIf base <> 0 Then
        Change = (last / base - 1#) * 100#
    End If
End Function

Private Sub ClearRow(ByRef m() As Variant, ByVal r As Long)
    Dim c As Long
    For c = 1 To MK_COLS
        m(r, c) = Empty
    Next c
End Sub

' ---------------------------------------------------------------- Yahoo

Private Function FetchYahoo(ByVal symEnc As String, ByRef d() As Double, _
                            ByRef v() As Double, ByRef n As Long) As Boolean
    On Error GoTo Fail
    Dim s As String
    s = HttpGetText("https://query1.finance.yahoo.com/v8/finance/chart/" & symEnc & "?range=2y&interval=1d", UA_BROWSER, "", "")
    If Len(s) = 0 Then Exit Function

    Dim pInd As Long
    pInd = InStr(1, s, """indicators""", vbBinaryCompare)
    If pInd = 0 Then Exit Function

    Dim tsArr() As String, clArr() As String
    If Not ArrayAfter(s, """timestamp"":[", 1, tsArr) Then Exit Function

    Dim ok As Boolean
    ok = ArrayAfter(s, """adjclose"":[{""adjclose"":[", pInd, clArr)
    If Not ok Then ok = ArrayAfter(s, """close"":[", pInd, clArr)
    If Not ok Then Exit Function

    Dim gmt As Double
    gmt = 0#
    Dim pG As Long
    pG = InStr(1, s, """gmtoffset"":", vbBinaryCompare)
    If pG > 0 Then gmt = NumAfterKey(s, pG + 12)

    Dim cnt As Long, i As Long
    cnt = UBound(tsArr)
    If UBound(clArr) < cnt Then cnt = UBound(clArr)
    ReDim d(1 To cnt + 1)
    ReDim v(1 To cnt + 1)
    Dim c As String
    For i = 0 To cnt
        c = Trim$(clArr(i))
        If Len(c) > 0 And c <> "null" Then
            n = n + 1
            d(n) = Fix((Val(tsArr(i)) + gmt) / 86400#)
            v(n) = Val(c)
        End If
    Next i
    FetchYahoo = (n >= 2)
    Exit Function
Fail:
    FetchYahoo = False
End Function

' Splits the JSON array that starts right after `marker` (searched from `startAt`).
Private Function ArrayAfter(ByVal s As String, ByVal marker As String, ByVal startAt As Long, _
                            ByRef outArr() As String) As Boolean
    Dim p As Long, q As Long
    p = InStr(startAt, s, marker, vbBinaryCompare)
    If p = 0 Then Exit Function
    p = p + Len(marker)
    q = InStr(p, s, "]", vbBinaryCompare)
    If q <= p Then Exit Function
    outArr = Split(Mid$(s, p, q - p), ",")
    ArrayAfter = True
End Function

' ---------------------------------------------------------------- HTTP

' Returns the response body, or "" on failure. Retries once (network error,
' 429, 5xx) unless the overall time budget is already spent. 4xx never retried.
Private Function HttpGetText(ByVal url As String, ByVal ua As String, _
                             ByVal referer As String, ByVal origin As String) As String
    Dim attempt As Long, http As Object, failNet As Boolean, st As Long, body As String
    For attempt = 1 To 2
        failNet = False
        st = 0
        body = ""
        On Error Resume Next
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        If http Is Nothing Then Exit Function
        http.SetTimeouts T_RESOLVE_MS, T_CONNECT_MS, T_SEND_MS, T_RECEIVE_MS
        http.Open "GET", url, False
        http.SetRequestHeader "User-Agent", ua
        http.SetRequestHeader "Accept", "application/json, text/csv, */*"
        If Len(referer) > 0 Then http.SetRequestHeader "Referer", referer
        If Len(origin) > 0 Then http.SetRequestHeader "Origin", origin
        http.Send
        If Err.Number <> 0 Then
            failNet = True
            Err.Clear
        Else
            st = http.Status
            If st = 200 Then body = http.ResponseText
            If Err.Number <> 0 Then body = "": Err.Clear
        End If
        Set http = Nothing
        On Error GoTo 0

        If st = 200 And Len(body) > 0 Then
            HttpGetText = body
            Exit Function
        End If
        If Not failNet And st <> 429 And st < 500 Then Exit Function      ' 4xx etc.: give up
        If Timer - mT0 > RETRY_BUDGET_SEC Or Timer < mT0 Then Exit Function
        If attempt = 1 Then SleepSec 1
    Next attempt
End Function

Private Sub SleepSec(ByVal sec As Double)
    Dim t As Double
    t = Timer
    Do While Timer - t < sec And Timer >= t
        DoEvents
    Loop
End Sub

' ---------------------------------------------------------------- parsing helpers

' Number following the colon after a key. `pos` points just past the key's
' closing quote (so the next non-space chars are ':' and the value).
Private Function NumAfterKey(ByVal s As String, ByVal pos As Long) As Double
    Dim i As Long, ch As String, buf As String
    i = pos
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = ":" Or ch = " " Then i = i + 1 Else Exit Do
    Loop
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If InStr("0123456789.-+eE", ch) > 0 Then
            buf = buf & ch
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    NumAfterKey = Val(buf)
End Function

' String value of "key": "value" found between positions lo..hi.
Private Function StrAfterKey(ByVal s As String, ByVal key As String, ByVal lo As Long, ByVal hi As Long) As String
    Dim p As Long, q As Long, r As Long
    p = InStr(lo, s, """" & key & """", vbBinaryCompare)
    If p = 0 Or p > hi Then Exit Function
    q = InStr(p + Len(key) + 2, s, """", vbBinaryCompare)
    If q = 0 Then Exit Function
    r = InStr(q + 1, s, """", vbBinaryCompare)
    If r = 0 Then Exit Function
    StrAfterKey = Mid$(s, q + 1, r - q - 1)
End Function

Private Function TitleCase(ByVal s As String) As String
    Dim w() As String, i As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    w = Split(s, " ")
    For i = 0 To UBound(w)
        If Len(w(i)) > 0 Then w(i) = UCase$(Left$(w(i), 1)) & LCase$(Mid$(w(i), 2))
    Next i
    TitleCase = Join(w, " ")
End Function

' ---------------------------------------------------------------- date helpers
' "Day number" here = whole days since 1970-01-01 (locale independent).

Private Function EpochDay() As Double
    EpochDay = CDbl(DateSerial(1970, 1, 1))
End Function

Private Function DayToStr(ByVal dayNo As Double) As String
    Dim dt As Date
    dt = CDate(dayNo + EpochDay())
    DayToStr = DateStr(dt)
End Function

Private Function DateStr(ByVal dt As Date) As String
    DateStr = Year(dt) & "-" & Right$("0" & Month(dt), 2) & "-" & Right$("0" & Day(dt), 2)
End Function

' "yyyy-mm-dd" -> date serial as Double (0 when unparsable)
Private Function DateValueIso(ByVal s As String) As Double
    On Error GoTo Bad
    If Len(s) < 10 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function
    DateValueIso = CDbl(DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2))))
    Exit Function
Bad:
    DateValueIso = 0#
End Function
