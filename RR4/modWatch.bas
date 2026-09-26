Attribute VB_Name = "modWatch"
Option Explicit

' ================================================================
'  WATCH (nav code W, W! = rebuild) - sheet "Watch"        2026-09-25
' ----------------------------------------------------------------
'  The WATCHLIST as a full worksheet. Until 2026-09-25 it was a 7-row
'  block on the RR4 page (B26:E35); that block is now a read-only
'  summary of this page's first 7 rows (PortfolioDashboard_v3.
'  ReadWatchlist -> WatchTop) and double-clicking it jumps here.
'
'  Data: sheet "WatchData", ListObject tblWatch, one row per name.
'    typed   TICKER | STRATEGY | ENTRY TGT | ADDED (default today)
'    cache   LAST | NOTES | CALL | STATUS | MOAT | RISK | CAT | R1 | R4 |
'            SIG | SIGDT     (program-written by W!, grey; the page is a
'            pure render of this table, so sorting / adding a name never
'            refetches anything - only W! does, like B! / V! / HC!)
'  Add by typing into the entry row on the page (TICKER + ENTRY TGT) or
'  under the table on WatchData; delete a name = delete its table row.
'
'  Library link: LibrarySummaryMap (modThesis) - note count, latest CALL
'  DATE (grey > 60 d, orange > 120 d, like the Library index), STATUS and
'  MOAT/RISK/CATALYST counts of the notes on that date.
'  Bias link: BiasSnapshot (modBias) - R1 / R4 as the Bias tables show
'  them, plus the latest SELL/BUY (short/long) rank+trend signal and date.
'  Double-click TICKER -> Library filtered to it; double-click any Bias
'  cell (R1 .. AGO) -> Bias page charts for it.
'
'  Sort (double-click a header): TICKER, ADDED/DAYS, DIST% - asc, desc,
'  off. Default = names at/below their ENTRY TGT first, then nearest to
'  target. The same order feeds the RR4 summary (WatchTop).
'
'  First run migrates the old RR4 block into tblWatch (SeedFromRR4).
'  Pure ASCII (VBE import rule).
' ================================================================

Public Const WATCH_PAGE As String = "Watch"
Public Const WATCH_DATA As String = "WatchData"
Public Const WATCH_TABLE As String = "tblWatch"

' --- tblWatch columns ---
Private Const WT_TICKER As Long = 1
Private Const WT_STRAT As Long = 2
Private Const WT_TGT As Long = 3
Private Const WT_ADDED As Long = 4
Private Const WT_LAST As Long = 5          ' first cache column
Private Const WT_NOTES As Long = 6
Private Const WT_CALL As Long = 7
Private Const WT_STATUS As Long = 8
Private Const WT_MOAT As Long = 9
Private Const WT_RISK As Long = 10
Private Const WT_CAT As Long = 11
Private Const WT_R1 As Long = 12
Private Const WT_R4 As Long = 13
Private Const WT_SIG As Long = 14
Private Const WT_SIGDT As Long = 15
Private Const WT_NCOL As Long = 15

' --- page rows (page coordinates; + NavOffset / NavLeft once the bar is on) ---
Private Const PG_TITLE As Long = 1
Private Const PG_LBL As Long = 2
Private Const PG_IN As Long = 3
Private Const PG_GRP As Long = 4
Private Const PG_HDR As Long = 5
Private Const PG_FIRST As Long = 6

' --- page columns ---
Private Const C_TK As Long = 1
Private Const C_MKT As Long = 2
Private Const C_STRAT As Long = 3
Private Const C_ADDED As Long = 4
Private Const C_DAYS As Long = 5
Private Const C_TGT As Long = 6
Private Const C_LAST As Long = 7
Private Const C_DIST As Long = 8
Private Const C_NOTES As Long = 9
Private Const C_CALL As Long = 10
Private Const C_STATUS As Long = 11
Private Const C_MRC As Long = 12
Private Const C_R1 As Long = 13
Private Const C_R4 As Long = 14
Private Const C_SIG As Long = 15
Private Const C_SIGDT As Long = 16
Private Const C_AGO As Long = 17
Private Const C_LASTCOL As Long = 17

Private Const FRESH_GREY As Long = 60
Private Const FRESH_ORANGE As Long = 120
Private Const SORT_NAME As String = "WATCHSORT"
Private Const FONT_FACE As String = "Consolas"
Private Const CLR_TEXT As Long = 14540253            ' RGB(221,221,221)
Private Const CLR_MUTED As Long = 8553090            ' RGB(130,130,130)

' ================================================================
'  Small helpers
' ================================================================
Private Function CellStr(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    CellStr = Trim(CStr(v))
End Function

Private Function NumOr0(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then NumOr0 = CDbl(v): Exit Function       ' IsNumeric(Date) is False
    If IsNumeric(v) Then NumOr0 = CDbl(v)
End Function

Private Function BareKey(ByVal s As String) As String
    BareKey = UCase(Trim(s))
    If Right(BareKey, 4) = ".TWO" Then BareKey = Left(BareKey, Len(BareKey) - 4)
    If Right(BareKey, 3) = ".TW" Then BareKey = Left(BareKey, Len(BareKey) - 3)
End Function

Private Function IsTWName(ByVal tk As String) As Boolean
    Dim u As String: u = UCase(Trim(tk))
    IsTWName = (InStr(u, ".TW") > 0) Or (u <> "" And IsNumeric(u))
End Function

Private Function NormTW(ByVal tk As String) As String
    Dim u As String: u = UCase(Trim(tk))
    If u = "" Then Exit Function
    If InStr(u, ".TW") > 0 Then
        NormTW = u
    ElseIf IsNumeric(u) Then
        NormTW = u & ".TW"
    End If
End Function

Private Function WatchTable() As ListObject
    On Error Resume Next
    Set WatchTable = ThisWorkbook.Worksheets(WATCH_DATA).ListObjects(WATCH_TABLE)
End Function

Private Function TableRows(ByVal lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    TableRows = lo.DataBodyRange.Rows.Count
End Function

Private Function PageSheet() As Worksheet
    On Error Resume Next
    Set PageSheet = ThisWorkbook.Worksheets(WATCH_PAGE)
End Function

Private Function SortMode() As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(SORT_NAME).RefersTo
    On Error GoTo 0
    s = Replace(s, "=", ""): s = Replace(s, """", "")
    SortMode = s
End Function

Private Sub SetSortMode(ByVal m As String)
    On Error Resume Next
    ThisWorkbook.Names(SORT_NAME).Delete
    On Error GoTo 0
    If m <> "" Then ThisWorkbook.Names.Add Name:=SORT_NAME, RefersTo:="=""" & m & """", Visible:=False
End Sub

' ================================================================
'  Data sheet + table
' ================================================================
Private Function EnsureWatchData() As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WATCH_DATA)
    On Error GoTo 0
    If ws Is Nothing Then
        Dim prev As Object: Set prev = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = WATCH_DATA
        ActiveWindow.DisplayGridlines = False
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If

    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = FONT_FACE: .Font.Size = 9: .Font.Color = CLR_TEXT
        .VerticalAlignment = xlCenter
    End With
    With ws.Cells(1, 1)
        .Value = "WATCH DATA  .  data for the Watch page (W) - one name per row: type TICKER / STRATEGY / ENTRY TGT (ADDED fills itself); grey columns are written by W!; delete a row to remove a name"
        .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With

    Dim lo As ListObject: Set lo = WatchTable()
    Dim created As Boolean
    Dim hdr As Variant
    hdr = Array("TICKER", "STRATEGY", "ENTRY TGT", "ADDED", "LAST", "NOTES", "CALL", "STATUS", "MOAT", "RISK", "CAT", "R1", "R4", "SIG", "SIGDT")
    Dim j As Long
    If lo Is Nothing Then
        For j = 0 To WT_NCOL - 1: ws.Cells(3, j + 1).Value = hdr(j): Next j
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(3, 1), ws.Cells(4, WT_NCOL)), , xlYes)
        lo.Name = WATCH_TABLE
        created = True
    End If
    lo.TableStyle = ""
    With lo.HeaderRowRange
        .Font.Color = RR4_ACCENT: .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RR4_LINE
    End With
    lo.ListColumns(WT_LAST).Range.Resize(1, WT_NCOL - WT_LAST + 1).Font.Color = CLR_MUTED
    Dim widths As Variant: widths = Array(12, 36, 12, 12, 11, 7, 12, 12, 6, 6, 6, 8, 8, 20, 12)
    For j = 0 To WT_NCOL - 1: ws.Columns(j + 1).ColumnWidth = widths(j): Next j
    If Not lo.DataBodyRange Is Nothing Then
        With lo.DataBodyRange
            .Interior.Color = RGB(8, 8, 8): .Font.Color = CLR_TEXT
        End With
        lo.ListColumns(WT_TICKER).DataBodyRange.NumberFormat = "@"
        lo.ListColumns(WT_STRAT).DataBodyRange.NumberFormat = "@"
        lo.ListColumns(WT_TGT).DataBodyRange.NumberFormat = "#,##0.00"
        lo.ListColumns(WT_ADDED).DataBodyRange.NumberFormat = "yyyy/mm/dd"
        lo.ListColumns(WT_LAST).DataBodyRange.NumberFormat = "#,##0.00"
        lo.ListColumns(WT_CALL).DataBodyRange.NumberFormat = "yyyy/mm/dd"
        lo.ListColumns(WT_R1).DataBodyRange.NumberFormat = "0.0"
        lo.ListColumns(WT_R4).DataBodyRange.NumberFormat = "0.0"
        lo.ListColumns(WT_SIGDT).DataBodyRange.NumberFormat = "yyyy/mm/dd"
        lo.ListColumns(WT_LAST).DataBodyRange.Resize(, WT_NCOL - WT_LAST + 1).Font.Color = CLR_MUTED
    End If
    Call WriteSheetCode(ws, "Option Explicit" & vbCrLf & vbCrLf & _
        "' WatchData (modWatch): editing the table redraws the Watch page" & vbCrLf & _
        "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    Call WatchDataChange(Me, Target)" & vbCrLf & _
        "End Sub" & vbCrLf, "WatchDataChange")

    If created Then Call SeedFromRR4(lo)
    Set EnsureWatchData = lo
End Function

' Public entry for PortfolioDashboard_v3.ReadWatchlist: make sure the table
' exists (seeding it from the RR4 block that is still intact at that point).
Public Sub WatchEnsure()
    If WatchTable() Is Nothing Then Call EnsureWatchData
End Sub

' One-off: copy the old RR4 WATCHLIST rows (B29:D35) into a fresh table.
Private Sub SeedFromRR4(ByVal lo As ListObject)
    Dim wsP As Worksheet
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets(NavSheetName("P"))
    On Error GoTo 0
    If wsP Is Nothing Then Exit Sub
    If Left(UCase(CellStr(wsP.Cells(RR4_WL_TITLE, RR4_LEFT + 1).Value)), 9) <> "WATCHLIST" Then Exit Sub
    Dim r As Long, tk As String
    For r = RR4_WL_FIRST To RR4_WL_LAST
        tk = UCase(CellStr(wsP.Cells(r, RR4_LEFT + 1).Value))
        If tk <> "" Then
            Call TableAppend(lo, tk, CellStr(wsP.Cells(r, RR4_LEFT + 2).Value), wsP.Cells(r, RR4_LEFT + 3).Value, Date)
        End If
    Next r
End Sub

Private Sub TableAppend(ByVal lo As ListObject, ByVal tk As String, ByVal st As String, _
                        ByVal tg As Variant, ByVal added As Date)
    Dim rowIdx As Long, r As Long
    If TableRows(lo) = 0 Then
        lo.Resize lo.Range.Resize(2)
    End If
    For r = 1 To lo.DataBodyRange.Rows.Count
        If CellStr(lo.DataBodyRange.Cells(r, WT_TICKER).Value) = "" Then rowIdx = r: Exit For
    Next r
    If rowIdx = 0 Then
        lo.Resize lo.Range.Resize(lo.Range.Rows.Count + 1)
        rowIdx = lo.DataBodyRange.Rows.Count
    End If
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    With lo.DataBodyRange
        .Cells(rowIdx, WT_TICKER).NumberFormat = "@": .Cells(rowIdx, WT_TICKER).Value = tk
        .Cells(rowIdx, WT_STRAT).NumberFormat = "@": .Cells(rowIdx, WT_STRAT).Value = st
        .Cells(rowIdx, WT_TGT).NumberFormat = "#,##0.00"
        If IsNumeric(tg) And Not IsEmpty(tg) Then .Cells(rowIdx, WT_TGT).Value = CDbl(tg) Else .Cells(rowIdx, WT_TGT).ClearContents
        .Cells(rowIdx, WT_ADDED).NumberFormat = "yyyy/mm/dd"
        .Cells(rowIdx, WT_ADDED).Value = added
        .Range(.Cells(rowIdx, WT_LAST), .Cells(rowIdx, WT_NCOL)).ClearContents
        .Rows(rowIdx).Interior.Color = RGB(8, 8, 8)
        .Range(.Cells(rowIdx, WT_TICKER), .Cells(rowIdx, WT_ADDED)).Font.Color = CLR_TEXT
        .Range(.Cells(rowIdx, WT_LAST), .Cells(rowIdx, WT_NCOL)).Font.Color = CLR_MUTED
    End With
    Application.EnableEvents = prevEv
End Sub

' ================================================================
'  Public readers for the other pages
' ================================================================

' Every ticker in the table (stored order, upper-cased, deduped), Empty if none.
Public Function WatchTickers() As Variant
    Dim lo As ListObject: Set lo = WatchTable()
    If TableRows(lo) = 0 Then Exit Function
    Dim v As Variant: v = lo.DataBodyRange.Value
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, tk As String
    For i = 1 To UBound(v, 1)
        tk = UCase(CellStr(v(i, WT_TICKER)))
        If tk <> "" Then seen(tk) = True
    Next i
    If seen.Count = 0 Then Exit Function
    WatchTickers = seen.keys
End Function

' First n names in page order as Variant(1..k, 1..3) = ticker / strategy /
' target (Empty when not a positive number); Empty if there are none. Feeds
' the RR4 page summary.
Public Function WatchTop(ByVal n As Long) As Variant
    Dim lo As ListObject: Set lo = WatchTable()
    If TableRows(lo) = 0 Then Exit Function
    Dim v As Variant: v = lo.DataBodyRange.Value
    Dim idx() As Long, cnt As Long
    cnt = OrderIdx(v, idx)
    If cnt = 0 Then Exit Function
    If cnt > n Then cnt = n
    Dim out() As Variant: ReDim out(1 To cnt, 1 To 3)
    Dim k As Long, i As Long, tg As Double
    For k = 1 To cnt
        i = idx(k)
        out(k, 1) = UCase(CellStr(v(i, WT_TICKER)))
        out(k, 2) = CellStr(v(i, WT_STRAT))
        tg = NumOr0(v(i, WT_TGT))
        If tg > 0 Then out(k, 3) = tg Else out(k, 3) = Empty
    Next k
    WatchTop = out
End Function

' Keeps the LAST cache current when the RR4 summary refreshes a price.
Public Sub WatchCacheLast(ByVal tk As String, ByVal px As Double)
    If px <= 0 Then Exit Sub
    Dim lo As ListObject: Set lo = WatchTable()
    If TableRows(lo) = 0 Then Exit Sub
    Dim r As Long, u As String: u = UCase(Trim(tk))
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    For r = 1 To lo.DataBodyRange.Rows.Count
        If UCase(CellStr(lo.DataBodyRange.Cells(r, WT_TICKER).Value)) = u Then
            lo.DataBodyRange.Cells(r, WT_LAST).Value = px
        End If
    Next r
    Application.EnableEvents = prevEv
End Sub

' ================================================================
'  Ordering
' ================================================================
Private Function DistOf(ByRef v As Variant, ByVal i As Long) As Double
    Dim lastPx As Double: lastPx = NumOr0(v(i, WT_LAST))
    Dim tg As Double: tg = NumOr0(v(i, WT_TGT))
    If lastPx > 0 And tg > 0 Then DistOf = (lastPx - tg) / tg Else DistOf = 1E+99
End Function

' True when row a should be listed before row b under the current sort mode.
Private Function RowBefore(ByRef v As Variant, ByVal a As Long, ByVal b As Long, ByVal mode As String) As Boolean
    Dim col As String, asc As Boolean
    If mode = "" Then
        col = "dflt": asc = True
    Else
        col = Split(mode, "|")(0)
        asc = (Split(mode, "|")(1) = "asc")
    End If
    Dim da As Double, db As Double, sa As String, sb As String
    Select Case col
        Case "ticker"
            sa = UCase(CellStr(v(a, WT_TICKER))): sb = UCase(CellStr(v(b, WT_TICKER)))
            If sa <> sb Then RowBefore = IIf(asc, sa < sb, sa > sb): Exit Function
        Case "added"
            da = NumOr0(v(a, WT_ADDED)): db = NumOr0(v(b, WT_ADDED))
            If da <> db Then RowBefore = IIf(asc, da < db, da > db): Exit Function
        Case "dist"
            da = DistOf(v, a): db = DistOf(v, b)
            If da <> db Then RowBefore = IIf(asc, da < db, da > db): Exit Function
        Case Else                                    ' default: at/below target first, then nearest
            da = DistOf(v, a): db = DistOf(v, b)
            Dim ha As Boolean, hb As Boolean
            ha = (da <= 0): hb = (db <= 0)
            If ha <> hb Then RowBefore = ha: Exit Function
            If da <> db Then RowBefore = (da < db): Exit Function
    End Select
    RowBefore = (a < b)                                 ' stable
End Function

' Row indexes (1-based into the table body) of the named rows, sorted.
Private Function OrderIdx(ByRef v As Variant, ByRef idx() As Long) As Long
    Dim n As Long: n = UBound(v, 1)
    ReDim idx(1 To n)
    Dim i As Long, m As Long
    For i = 1 To n
        If CellStr(v(i, WT_TICKER)) <> "" Then m = m + 1: idx(m) = i
    Next i
    If m = 0 Then Exit Function
    ReDim Preserve idx(1 To m)
    Dim mode As String: mode = SortMode()
    Dim j As Long, best As Long, t As Long
    For i = 1 To m - 1
        best = i
        For j = i + 1 To m
            If RowBefore(v, idx(j), idx(best), mode) Then best = j
        Next j
        If best <> i Then t = idx(i): idx(i) = idx(best): idx(best) = t
    Next i
    OrderIdx = m
End Function

' ================================================================
'  W! - refresh the cache, rebuild the page
' ================================================================
Public Sub BuildWatchPage()
    Dim lo As ListObject: Set lo = EnsureWatchData()
    Dim ws As Worksheet: Set ws = EnsurePageSheet()
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fail

    Dim n As Long
    If TableRows(lo) > 0 Then n = RefreshCache(lo)
    Call NavStrip(ws)
    Call DrawShell(ws)
    Call DrawWatchRows(ws)
    Call FinishPage(ws)

    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
    Call NavGoto("W", ws)
    Call NavNotify("WATCH done - " & n & " names refreshed")
    Exit Sub

Fail:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    Call NavNotify("WATCH error: " & msg, True)
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
End Sub

' Fetches price + Library + Bias for every named row and writes the cache
' columns back in one block. Returns the number of names processed.
Private Function RefreshCache(ByVal lo As ListObject) As Long
    Dim v As Variant: v = lo.DataBodyRange.Value
    Dim n As Long: n = UBound(v, 1)
    Dim i As Long, tk As String

    ' one batched Taiwan price prefetch, like UP does
    Dim tw() As String, nt As Long
    ReDim tw(0 To n)
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        tk = NormTW(CellStr(v(i, WT_TICKER)))
        If tk <> "" And Not seen.Exists(tk) Then seen(tk) = 1: tw(nt) = tk: nt = nt + 1
    Next i
    If nt > 0 Then
        ReDim Preserve tw(0 To nt - 1)
        On Error Resume Next
        modMISPrice.PrefetchMISPrices tw
        On Error GoTo 0
    End If

    Dim libMap As Object: Set libMap = LibrarySummaryMap()
    Dim px As Double, a As Variant, key As String
    Dim r1 As Double, r4 As Double, hasR4 As Boolean, sigTxt As String, sigDt As Date, ok As Boolean
    Dim done As Long
    For i = 1 To n
        tk = UCase(CellStr(v(i, WT_TICKER)))
        If tk <> "" Then
            done = done + 1
            Call NavNotify("WATCH " & done & " " & tk & " ...")
            px = 0
            Dim tkArg As String: tkArg = tk
            On Error Resume Next
            px = GetStockPrice(tkArg)
            Err.Clear
            On Error GoTo 0
            If px > 0 Then v(i, WT_LAST) = px

            key = BareKey(tk)
            If libMap.Exists(key) Then
                a = libMap(key)
                v(i, WT_NOTES) = a(0)
                If a(1) > 0 Then v(i, WT_CALL) = a(1) Else v(i, WT_CALL) = Empty
                v(i, WT_STATUS) = a(2)
                v(i, WT_MOAT) = a(3): v(i, WT_RISK) = a(4): v(i, WT_CAT) = a(5)
            Else
                v(i, WT_NOTES) = 0
                v(i, WT_CALL) = Empty: v(i, WT_STATUS) = Empty
                v(i, WT_MOAT) = Empty: v(i, WT_RISK) = Empty: v(i, WT_CAT) = Empty
            End If

            ok = False
            On Error Resume Next
            ok = BiasSnapshot(tk, r1, hasR4, r4, sigTxt, sigDt)
            Err.Clear
            On Error GoTo 0
            If ok Then
                v(i, WT_R1) = r1
                If hasR4 Then v(i, WT_R4) = r4 Else v(i, WT_R4) = "-"
                v(i, WT_SIG) = sigTxt
                If sigTxt <> "" Then v(i, WT_SIGDT) = sigDt Else v(i, WT_SIGDT) = Empty
            Else
                v(i, WT_R1) = "n/a": v(i, WT_R4) = Empty: v(i, WT_SIG) = Empty: v(i, WT_SIGDT) = Empty
            End If
        End If
    Next i

    Dim nc As Long: nc = WT_NCOL - WT_LAST + 1
    Dim outv() As Variant: ReDim outv(1 To n, 1 To nc)
    Dim j As Long
    For i = 1 To n
        For j = 1 To nc
            outv(i, j) = v(i, WT_LAST + j - 1)
        Next j
    Next i
    lo.DataBodyRange.Columns(WT_LAST).Resize(, nc).Value = outv
    RefreshCache = done
End Function

' ================================================================
'  Page sheet
' ================================================================
Private Function EnsurePageSheet() As Worksheet
    Dim ws As Worksheet: Set ws = PageSheet()
    If ws Is Nothing Then
        Dim prev As Object: Set prev = ActiveSheet
        Dim after As Object
        On Error Resume Next
        Set after = ThisWorkbook.Worksheets("Bias")
        On Error GoTo 0
        If after Is Nothing Then Set after = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set ws = ThisWorkbook.Worksheets.Add(After:=after)
        ws.Name = WATCH_PAGE
        ActiveWindow.DisplayGridlines = False
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If
    Set EnsurePageSheet = ws
End Function

Private Sub DrawShell(ByVal ws As Worksheet)
    ws.Cells.Clear
    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = FONT_FACE
        .Font.Size = 9
        .Font.Color = CLR_TEXT
        .VerticalAlignment = xlCenter
        .RowHeight = 16
    End With
    With ws.Cells(PG_TITLE, 1)
        .Value = "WATCHLIST"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 14
    End With
    ws.Rows(PG_TITLE).RowHeight = 24

    Call StackLabel(ws.Cells(PG_LBL, C_TK), "TICKER")
    Call StackLabel(ws.Cells(PG_LBL, C_STRAT), "STRATEGY")
    Call StackLabel(ws.Cells(PG_LBL, C_TGT), "ENTRY TGT")
    ws.Rows(PG_LBL).RowHeight = 18
    ws.Rows(PG_IN).RowHeight = 20
    Call InputCell(ws.Cells(PG_IN, C_TK), "@", xlCenter)
    Call InputCell(ws.Cells(PG_IN, C_STRAT), "@", xlLeft)
    Call InputCell(ws.Cells(PG_IN, C_TGT), "#,##0.00", xlCenter)
    Dim c As Long
    For c = C_TK To C_TGT - 1
        With ws.Cells(PG_IN, c).Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Color = RGB(0, 0, 0): .Weight = xlMedium
        End With
    Next c

    ' group bands
    Dim g As Variant, gc As Variant, k As Long
    g = Array("WATCH", "LIBRARY", "BIAS  (R1 = EMA20 short, R4 = EMA200 long)")
    gc = Array(C_TK, C_NOTES, C_R1)
    For k = 0 To 2
        With ws.Cells(PG_GRP, gc(k))
            .Value = g(k)
            .Font.Color = RR4_ACCENT: .Font.Bold = True
        End With
    Next k
    Dim widths As Variant
    widths = Array(12, 6, 34, 11, 6, 11, 11, 9, 7, 12, 12, 8, 8, 8, 26, 11, 6)
    For c = 0 To C_LASTCOL - 1: ws.Columns(c + 1).ColumnWidth = widths(c): Next c
    Call DrawHeader(ws, 0, 0)
End Sub

Private Sub DrawHeader(ByVal ws As Worksheet, ByVal off As Long, ByVal lc As Long)
    Dim hdr As Variant
    hdr = Array("TICKER", "MKT", "STRATEGY", "ADDED", "DAYS", "ENTRY TGT", "LAST", "DIST%", _
                "NOTES", "LAST CALL", "STATUS", "M/R/C", "R1", "R4", "SIGNAL", "SIG DATE", "AGO")
    Dim mode As String: mode = SortMode()
    Dim c As Long, txt As String, arrow As String
    For c = 1 To C_LASTCOL
        txt = hdr(c - 1)
        arrow = ""
        If mode <> "" Then
            Dim sc As String, sd As String
            sc = Split(mode, "|")(0): sd = Split(mode, "|")(1)
            If (sc = "ticker" And c = C_TK) Or (sc = "added" And (c = C_ADDED Or c = C_DAYS)) Or (sc = "dist" And c = C_DIST) Then
                arrow = IIf(sd = "asc", " " & ChrW(&H25B2), " " & ChrW(&H25BC))
            End If
        End If
        With ws.Cells(PG_HDR + off, c + lc)
            .Value = txt & arrow
            .Font.Color = RGB(0, 200, 255)
            .Font.Size = 9
            .HorizontalAlignment = IIf(c = C_STRAT Or c = C_SIG, xlLeft, xlCenter)
        End With
    Next c
    With ws.Range(ws.Cells(PG_HDR + off, 1 + lc), ws.Cells(PG_HDR + off, C_LASTCOL + lc)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With
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

Private Sub InputCell(ByVal cell As Range, ByVal fmt As String, ByVal align As Long)
    With cell
        .NumberFormat = fmt
        .Value = ""
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = align
    End With
End Sub

' Clears the data rows (below the header) and re-blacks them.
Private Sub ClearRows(ByVal ws As Worksheet, ByVal r0 As Long, ByVal lc As Long)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, C_TK + lc).End(xlUp).Row + 5
    If lastRow < r0 + 60 Then lastRow = r0 + 60
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r0, 1), ws.Cells(lastRow, C_LASTCOL + lc + 2))
    rng.ClearContents
    rng.Interior.Color = RGB(0, 0, 0)
    rng.Font.Bold = False
    rng.Font.Color = CLR_TEXT
    rng.Borders.LineStyle = xlNone
End Sub

Private Function StatusGood(ByVal st As String) As Boolean
    Select Case LCase(Trim(st))
        Case "robust", "solid", "growing": StatusGood = True
    End Select
End Function

' Renders tblWatch (its cache columns) in the current sort order.
Private Sub DrawWatchRows(ByVal ws As Worksheet)
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim r0 As Long: r0 = PG_FIRST + off
    Call ClearRows(ws, r0, lc)
    Call DrawHeader(ws, off, lc)

    Dim lo As ListObject: Set lo = WatchTable()
    Dim cnt As Long, k As Long, v As Variant, idx() As Long
    If TableRows(lo) > 0 Then
        v = lo.DataBodyRange.Value
        cnt = OrderIdx(v, idx)
    End If
    If cnt = 0 Then
        With ws.Cells(r0, 1 + lc)
            .Value = "(empty - type TICKER + ENTRY TGT in the boxes above, then press Enter)"
            .Font.Color = CLR_MUTED
        End With
        Exit Sub
    End If
    For k = 1 To cnt
        Call DrawOneRow(ws, r0 + k - 1, lc, v, idx(k))
    Next k
    With ws.Range(ws.Cells(r0 + cnt - 1, 1 + lc), ws.Cells(r0 + cnt - 1, C_LASTCOL + lc)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With
    Call PaintDeleteHandles(ws)
End Sub

' 2026-09-26: an x in the blank column A of every data row - double-click it to
' remove that name (WatchPageDoubleClick -> DeleteWatchRow). Only meaningful
' once the bar's blank column exists (NavLeft >= 1): during BuildWatchPage the
' rows are drawn on the stripped page, so FinishPage calls this again after NavAdd.
Private Sub PaintDeleteHandles(ByVal ws As Worksheet)
    Dim lc As Long: lc = NavLeft(ws)
    If lc < 1 Then Exit Sub
    Dim r As Long, t As String
    For r = PG_FIRST + NavOffset(ws) To PG_FIRST + NavOffset(ws) + 1000
        t = CellStr(ws.Cells(r, C_TK + lc).Value)
        If t = "" Or Left$(t, 1) = "(" Then Exit For
        With ws.Cells(r, 1)
            .Value = ChrW(&HD7)
            .HorizontalAlignment = xlCenter
            .Font.Color = CLR_MUTED
        End With
    Next r
End Sub

Private Sub DrawOneRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lc As Long, ByRef v As Variant, ByVal i As Long)
    Dim tk As String: tk = UCase(CellStr(v(i, WT_TICKER)))
    Dim tg As Double: tg = NumOr0(v(i, WT_TGT))
    Dim lastPx As Double: lastPx = NumOr0(v(i, WT_LAST))
    Dim hit As Boolean: hit = (lastPx > 0 And tg > 0 And lastPx <= tg)

    With ws.Cells(r, C_TK + lc)
        .NumberFormat = "@": .Value = tk
        .HorizontalAlignment = xlCenter
        .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With
    On Error Resume Next
    ws.Cells(r, C_TK + lc).Errors(xlNumberAsText).Ignore = True      ' "1303" typed as text
    ws.Cells(r, C_MRC + lc).Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
    With ws.Cells(r, C_MKT + lc)
        .Value = IIf(IsTWName(tk), "TW", "US")
        .HorizontalAlignment = xlCenter: .Font.Color = CLR_MUTED
    End With
    With ws.Cells(r, C_STRAT + lc)
        .NumberFormat = "@": .Value = CellStr(v(i, WT_STRAT))
        .HorizontalAlignment = xlLeft
        .Font.Name = "Noto Sans TC"
    End With
    Dim added As Double: added = NumOr0(v(i, WT_ADDED))
    With ws.Cells(r, C_ADDED + lc)
        If added > 0 Then .NumberFormat = "yyyy/mm/dd": .Value = CDate(added)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(r, C_DAYS + lc)
        If added > 0 Then .Value = CLng(Date - CDate(added))
        .HorizontalAlignment = xlCenter: .Font.Color = CLR_MUTED
    End With
    With ws.Cells(r, C_TGT + lc)
        .NumberFormat = "#,##0.00"
        If tg > 0 Then .Value = tg
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(r, C_LAST + lc)
        .NumberFormat = "#,##0.00"
        If lastPx > 0 Then .Value = lastPx Else .Value = "-"
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(r, C_DIST + lc)
        If lastPx > 0 And tg > 0 Then
            .NumberFormat = "+0.0%;-0.0%"
            .Value = (lastPx - tg) / tg
        Else
            .Value = "-"
        End If
        .HorizontalAlignment = xlCenter
    End With
    If hit Then
        With ws.Range(ws.Cells(r, C_TK + lc), ws.Cells(r, C_DIST + lc))
            .Interior.Color = RGB(60, 30, 0)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
        End With
    End If

    ' --- Library ---
    Dim notesV As Variant: notesV = v(i, WT_NOTES)
    With ws.Cells(r, C_NOTES + lc)
        If IsEmpty(notesV) Then
            .Value = "-": .Font.Color = CLR_MUTED
        Else
            .Value = notesV
            .Font.Color = IIf(NumOr0(notesV) = 0, CLR_MUTED, CLR_TEXT)
        End If
        .HorizontalAlignment = xlCenter
    End With
    Dim callD As Double: callD = NumOr0(v(i, WT_CALL))
    With ws.Cells(r, C_CALL + lc)
        .HorizontalAlignment = xlCenter
        If callD > 0 Then
            .NumberFormat = "yyyy/mm/dd": .Value = CDate(callD)
            Dim age As Long: age = CLng(Date - CDate(callD))
            If age > FRESH_ORANGE Then
                .Font.Color = RR4_ACCENT: .Font.Bold = True
            ElseIf age > FRESH_GREY Then
                .Font.Color = CLR_MUTED
            End If
        Else
            .Value = "-": .Font.Color = CLR_MUTED
        End If
    End With
    Dim st As String: st = CellStr(v(i, WT_STATUS))
    With ws.Cells(r, C_STATUS + lc)
        .HorizontalAlignment = xlCenter
        If st <> "" Then
            .Value = st
            .Font.Color = IIf(StatusGood(st), RGB(60, 200, 90), RGB(220, 60, 60))
        Else
            .Value = "-": .Font.Color = CLR_MUTED
        End If
    End With
    With ws.Cells(r, C_MRC + lc)
        .HorizontalAlignment = xlCenter
        If IsEmpty(v(i, WT_MOAT)) And IsEmpty(v(i, WT_RISK)) And IsEmpty(v(i, WT_CAT)) Then
            .Value = "-": .Font.Color = CLR_MUTED
        Else
            .NumberFormat = "@"
            .Value = CLng(NumOr0(v(i, WT_MOAT))) & "/" & CLng(NumOr0(v(i, WT_RISK))) & "/" & CLng(NumOr0(v(i, WT_CAT)))
        End If
    End With

    ' --- Bias ---
    Call DrawHeat(ws.Cells(r, C_R1 + lc), v(i, WT_R1))
    Call DrawHeat(ws.Cells(r, C_R4 + lc), v(i, WT_R4))
    Dim sig As String: sig = CellStr(v(i, WT_SIG))
    With ws.Cells(r, C_SIG + lc)
        .HorizontalAlignment = xlLeft
        If sig <> "" Then
            .Value = sig
            .Font.Color = IIf(InStr(sig, "SELL") > 0, RGB(220, 60, 60), RGB(60, 200, 90))
            .Font.Bold = True
        ElseIf IsEmpty(v(i, WT_R1)) Then
            .Value = "-": .Font.Color = CLR_MUTED
        Else
            .Value = "none": .Font.Color = CLR_MUTED
        End If
    End With
    Dim sd As Double: sd = NumOr0(v(i, WT_SIGDT))
    With ws.Cells(r, C_SIGDT + lc)
        .HorizontalAlignment = xlCenter
        If sd > 0 Then .NumberFormat = "yyyy/mm/dd": .Value = CDate(sd) Else .Value = "-": .Font.Color = CLR_MUTED
    End With
    With ws.Cells(r, C_AGO + lc)
        .HorizontalAlignment = xlCenter
        If sd > 0 Then .Value = CLng(Date - CDate(sd)) Else .Value = "-": .Font.Color = CLR_MUTED
    End With
End Sub

Private Sub DrawHeat(ByVal cell As Range, ByVal v As Variant)
    With cell
        .HorizontalAlignment = xlCenter
        If IsEmpty(v) Then
            .Value = "-": .Font.Color = CLR_MUTED
        ElseIf IsNumeric(v) And VarType(v) <> vbString Then
            .NumberFormat = "0.0"
            .Value = CDbl(v)
            .Interior.Color = CorrHeatBg(CDbl(v) / 100)
            .Font.Color = CorrHeatFg(CDbl(v) / 100)
            .Font.Bold = True
        Else
            .Value = CStr(v): .Font.Color = CLR_MUTED
        End If
    End With
End Sub

Private Sub FinishPage(ByVal ws As Worksheet)
    On Error Resume Next
    Call NavAdd(ws, "W")
    Call PaintDeleteHandles(ws)
    Call EnsureSheetCode(ws)
    On Error GoTo 0
End Sub

' ================================================================
'  Events (called from the sheet code written by EnsureSheetCode)
' ================================================================

' Page: the entry row (TICKER / STRATEGY / ENTRY TGT).
Public Sub WatchPageChange(ByVal ws As Worksheet, ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    If Target.Row <> PG_IN + off Then Exit Sub
    If Target.Column <> C_TK + lc And Target.Column <> C_STRAT + lc And Target.Column <> C_TGT + lc Then Exit Sub
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fin
    Call CommitEntry(ws, off, lc)
Fin:
    Application.EnableEvents = prevEv
    If Err.Number <> 0 Then Call NavNotify("WATCH entry error: " & Err.Description, True)
End Sub

' Commits once TICKER and a positive ENTRY TGT are both in. An existing
' ticker is updated (strategy / target) instead of duplicated.
Private Sub CommitEntry(ByVal ws As Worksheet, ByVal off As Long, ByVal lc As Long)
    Dim tk As String: tk = UCase(CellStr(ws.Cells(PG_IN + off, C_TK + lc).Value))
    Dim st As String: st = CellStr(ws.Cells(PG_IN + off, C_STRAT + lc).Value)
    Dim tgv As Variant: tgv = ws.Cells(PG_IN + off, C_TGT + lc).Value
    If tk = "" Then Exit Sub
    If IsEmpty(tgv) Then Exit Sub
    If Not IsNumeric(tgv) Then Exit Sub
    If CDbl(tgv) <= 0 Then Exit Sub

    Dim lo As ListObject: Set lo = WatchTable()
    If lo Is Nothing Then Set lo = EnsureWatchData()
    Dim r As Long, found As Long
    If TableRows(lo) > 0 Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            If UCase(CellStr(lo.DataBodyRange.Cells(r, WT_TICKER).Value)) = tk Then found = r: Exit For
        Next r
    End If
    If found > 0 Then
        lo.DataBodyRange.Cells(found, WT_TGT).Value = CDbl(tgv)
        If st <> "" Then lo.DataBodyRange.Cells(found, WT_STRAT).Value = st
        Call NavNotify("WATCH " & tk & " updated @ " & Format(CDbl(tgv), "#,##0.00"))
    Else
        Call TableAppend(lo, tk, st, CDbl(tgv), Date)
        Call NavNotify("WATCH + " & tk & " @ " & Format(CDbl(tgv), "#,##0.00") & "  (W! fetches its Library / Bias columns)")
    End If
    ws.Cells(PG_IN + off, C_TK + lc).Value = ""
    ws.Cells(PG_IN + off, C_STRAT + lc).Value = ""
    ws.Cells(PG_IN + off, C_TGT + lc).Value = ""
    Call DrawWatchRows(ws)
    ws.Cells(PG_IN + off, C_TK + lc).Select
End Sub

' Data sheet: keep TICKER upper-case, default ADDED, redraw the page.
Public Sub WatchDataChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(WATCH_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.Range) Is Nothing Then Exit Sub
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    Dim body As Range: Set body = Intersect(Target, lo.DataBodyRange)
    If Not body Is Nothing Then
        Dim c As Range, rr As Long
        For Each c In body.Rows
            rr = c.Row - lo.HeaderRowRange.Row
            If rr >= 1 And rr <= lo.ListRows.Count Then
                With lo.DataBodyRange
                    If CellStr(.Cells(rr, WT_TICKER).Value) <> "" Then
                        .Cells(rr, WT_TICKER).NumberFormat = "@"
                        .Cells(rr, WT_TICKER).Value = UCase(CellStr(.Cells(rr, WT_TICKER).Value))
                        If CellStr(.Cells(rr, WT_ADDED).Value) = "" Then
                            .Cells(rr, WT_ADDED).NumberFormat = "yyyy/mm/dd"
                            .Cells(rr, WT_ADDED).Value = Date
                        End If
                    End If
                End With
            End If
        Next c
    End If
    Dim wp As Worksheet: Set wp = PageSheet()
    If Not wp Is Nothing Then
        If NavPageCode(wp) = "W" Then Call DrawWatchRows(wp)
    End If
    On Error GoTo 0
    Application.EnableEvents = prevEv
End Sub

' Page double-click: header = cycle a sort; TICKER -> Library; Bias cells -> Bias.
Public Sub WatchPageDoubleClick(ByVal ws As Worksheet, ByVal Target As Range, ByRef Cancel As Boolean)
    Dim t As Range: Set t = Target.Cells(1, 1)
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    ' 2026-09-26: the x in the blank column A of a data row removes that name
    If lc >= 1 And t.Column = 1 And t.Row >= PG_FIRST + off Then
        If CellStr(t.Value) = ChrW(&HD7) Then
            Cancel = True
            Call DeleteWatchRow(ws, t.Row, lc)
        End If
        Exit Sub
    End If
    Dim c As Long: c = t.Column - lc
    If c < 1 Or c > C_LASTCOL Then Exit Sub

    If t.Row = PG_HDR + off Then
        Dim key As String, first As String
        Select Case c
            Case C_TK: key = "ticker": first = "asc"
            Case C_ADDED, C_DAYS: key = "added": first = "desc"
            Case C_DIST: key = "dist": first = "desc"
        End Select
        If key = "" Then Exit Sub
        Cancel = True
        Dim cur As String: cur = SortMode()
        Dim other As String: other = IIf(first = "asc", "desc", "asc")
        If cur = key & "|" & first Then
            Call SetSortMode(key & "|" & other)
        ElseIf cur = key & "|" & other Then
            Call SetSortMode("")
        Else
            Call SetSortMode(key & "|" & first)
        End If
        Dim prevEv As Boolean: prevEv = Application.EnableEvents
        Application.EnableEvents = False
        Call DrawWatchRows(ws)
        Application.EnableEvents = prevEv
        Exit Sub
    End If

    If t.Row < PG_FIRST + off Then Exit Sub
    Dim tk As String: tk = UCase(CellStr(ws.Cells(t.Row, C_TK + lc).Value))
    If tk = "" Then Exit Sub
    If c = C_TK Then
        Cancel = True
        If LibraryOpenTicker(tk) Then
            Call NavGoto("L", ws)
        Else
            Call NavNotify("Library is not built yet - run L!", True)
        End If
    ElseIf c >= C_R1 Then
        Cancel = True
        If BiasOpenTicker(tk) Then
            Call NavGoto("B", ws)
        Else
            Call NavNotify("Bias is not built yet - run B!", True)
        End If
    End If
End Sub

' Double-click on the x in column A: confirm, delete that name's row from
' tblWatch, redraw the Watch page only (no price / Bias refetch - the other
' names keep their cached values). The RR4 WATCHLIST summary catches up on the
' next UP. Nothing is kept: same as deleting the row on WatchData by hand.
Private Sub DeleteWatchRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lc As Long)
    Dim tk As String: tk = UCase(CellStr(ws.Cells(r, C_TK + lc).Value))
    If tk = "" Then Exit Sub
    Dim st As String: st = CellStr(ws.Cells(r, C_STRAT + lc).Value)
    If MsgBox("Remove " & tk & IIf(st <> "", "  (" & st & ")", "") & " from the watchlist?", _
              vbYesNo + vbQuestion + vbDefaultButton2, "WATCHLIST") <> vbYes Then Exit Sub
    Dim lo As ListObject: Set lo = WatchTable()
    If lo Is Nothing Then Exit Sub
    If TableRows(lo) = 0 Then Exit Sub
    Dim v As Variant: v = lo.DataBodyRange.Value
    Dim i As Long, hit As Long
    For i = 1 To UBound(v, 1)
        If UCase(CellStr(v(i, WT_TICKER))) = tk Then hit = i: Exit For
    Next i
    If hit = 0 Then
        Call NavNotify("WATCHLIST: " & tk & " is not in tblWatch - run W!", True)
        Exit Sub
    End If
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    lo.ListRows(hit).Delete
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    On Error GoTo 0
    If errN = 0 Then Call DrawWatchRows(ws)
    Application.EnableEvents = prevEv
    If errN = 0 Then
        Call NavNotify("WATCHLIST: removed " & tk)
    Else
        Call NavNotify("WATCHLIST: delete failed - " & errD, True)
    End If
End Sub

' RR4 page summary double-click: open the Watch page on that name.
Public Sub WatchGoto(ByVal src As Worksheet, ByVal tk As String)
    Dim ws As Worksheet: Set ws = PageSheet()
    If ws Is Nothing Then
        Call NavStatus(src, "[W] Watch is not built yet - run W!", True)
        Exit Sub
    End If
    Call NavGoto("W", src)
    tk = UCase(Trim(tk))
    If tk = "" Then Exit Sub
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim r As Long
    For r = PG_FIRST + off To PG_FIRST + off + 500
        If UCase(CellStr(ws.Cells(r, C_TK + lc).Value)) = tk Then
            ws.Cells(r, C_TK + lc).Select
            Exit For
        End If
        If CellStr(ws.Cells(r, C_TK + lc).Value) = "" Then Exit For
    Next r
End Sub

' ================================================================
'  Sheet event code (document modules are not exported as .bas)
' ================================================================
Private Sub EnsureSheetCode(ByVal ws As Worksheet)
    Call WriteSheetCode(ws, "Option Explicit" & vbCrLf & vbCrLf & _
        "' Watch page (modWatch): entry row + double-click actions" & vbCrLf & _
        "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    Call WatchPageChange(Me, Target)" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & vbCrLf & _
        "    Call WatchPageDoubleClick(Me, Target, Cancel)" & vbCrLf & _
        "End Sub" & vbCrLf, "WatchPageDoubleClick")
End Sub

Private Sub WriteSheetCode(ByVal ws As Worksheet, ByVal code As String, ByVal marker As String)
    On Error GoTo Skip
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 Then
            If comp.Properties("Name").Value = ws.Name Then
                Dim cm As Object: Set cm = comp.CodeModule
                If cm.CountOfLines > 0 Then
                    If InStr(cm.Lines(1, cm.CountOfLines), marker) > 0 Then Exit Sub
                    cm.DeleteLines 1, cm.CountOfLines
                End If
                cm.AddFromString code
                Exit Sub
            End If
        End If
    Next comp
    Exit Sub
Skip:
    Call NavNotify(ws.Name & " built, but its sheet event code could not be written (enable Trust access to the VBA project object model, or paste RR4/SheetWatch_Code.txt by hand)", True)
End Sub
