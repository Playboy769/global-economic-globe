Attribute VB_Name = "modThesis"
Option Explicit

' ================================================================
'  LIBRARY v2 - thesis notes workbench (nav code L, L! = rebuild)
'  (sheet + code renamed 2026-09-13 late: "Thesis Library" / TH -> "Library" / L)
' ----------------------------------------------------------------
'  2026-09-13 (evening).  v1 kept one row per 500-word thesis; v2 keeps one
'  row per THEME note and shows them grouped by ticker. Nothing is drawn until
'  QUERY or KEYWORD is typed; KEYWORD alone searches each ticker's latest call,
'  a QUERY shows every call of the queried tickers, newest first:
'
'    ThesisNotes (data sheet)  ListObject tblNotes, one note per row:
'      TARGET | TYPE (stock / macro) | CALL DATE | STATUS | ROLE |
'      THEME | BEHAVIOR | T0 Q&A | T1 CALL | T1L FILING | T2 RESEARCH |
'      T3 MEDIA | D OWN
'      (2026-09-15: the single EVIDENCE column became six evidence TIERS,
'      one note's evidence sits in exactly one of them:
'        T0  earnings-call Q&A (unscripted answers, refusals, dodges)
'        T1  the company's scripted statements (prepared remarks, guidance,
'            decks, press releases)
'        T1L statutory filings and official data - lag information
'            (annual report, 10-Q/K, monthly sales, MOPS, official stats)
'        T2  third-party research (sell side, buy side, consensus, industry
'            research houses, data sites)
'        T3  media, rumours, unattributed reports
'        D   own derivations (cross-check tables, valuation page, report
'            tabs, framework claims) - drawn as a Greek delta
'      EnsureNotesSheet upgrades an 8-column table in place: EVIDENCE is
'      renamed T1L FILING and the other five tiers are inserted around it.)
'      STATUS  Robust Solid Growing (green) / Slowing Sluggish Challenging
'              Contraction Warning (red)
'      ROLE    MOAT / RISK / CATALYST / blank
'      Typing there redraws the view (TYPE defaults to stock, CALL DATE to
'      today).  Delete a note = delete its table row there.
'
'    Thesis Library (view, carries the nav bar)
'      page row 1  title
'      row 2/3     QUERY (tickers, comma = several, blank = all) and KEYWORD
'                  (searched in THEME / BEHAVIOR / all six tiers, comma = OR)
'                  labels above their inputs; "N calls drawn . M notes"
'      row 5       totals over all data (notes . calls . tickers)
'      row 7       header; double-click its TICKER cell = sort blocks by
'                  latest call date (newest first) <-> ticker A-Z (THSORT)
'      row 8+      STOCK section, then MACRO (THEME | BEHAVIOR | T0 | T1 |
'                  T1L | T2 | T3 | D side by side): one block per target with the
'                  target bold and its latest call date under it, one line
'                  per note; notes of older calls follow, each call's date
'                  on its first row.  Double-click a target -> the
'                  reading panel shows its archived six-section thesis.
'
'    ThesisArchive (hidden)  the v1 tblThesis, untouched: the first TH! run
'      after this change renames the old "Thesis Library" sheet to it and
'      strips its v1 event code.
'
'  Chinese labels are ChrW-built so this file stays ASCII (VBE import rule).
' ================================================================

Public Const THESIS_SHEET As String = "Library"
Public Const THESIS_SHEET_OLD As String = "Thesis Library"   ' v1 / early-v2 tab name, migrated by EnsureViewSheet
Public Const NOTES_SHEET As String = "ThesisNotes"
Public Const NOTES_TABLE As String = "tblNotes"
Public Const ARCHIVE_SHEET As String = "ThesisArchive"
Public Const THESIS_TABLE As String = "tblThesis"        ' the archived v1 table

Private Const NT_TARGET As Long = 1
Private Const NT_TYPE As Long = 2
Private Const NT_DATE As Long = 3
Private Const NT_STATUS As Long = 4
Private Const NT_ROLE As Long = 5
Private Const NT_THEME As Long = 6
Private Const NT_BEHAV As Long = 7
Private Const NT_T0 As Long = 8                          ' first tier column; tiers are NT_T0 .. NT_T0 + NTIER - 1
Private Const NTIER As Long = 6
Private Const NT_NCOL As Long = 13

' view geometry (page coordinates; the bar adds NavOffset rows / NavLeft columns)
Private Const PG_TITLE As Long = 1
Private Const PG_LBL As Long = 2
Private Const PG_IN As Long = 3
Private Const PG_TOTAL As Long = 5
Private Const PG_HDR As Long = 7
Private Const PG_LIST As Long = 8
Private Const C_TGT As Long = 1
Private Const C_STATUS As Long = 2
Private Const C_ROLE As Long = 3
Private Const C_THEME As Long = 4
Private Const C_BEHAV As Long = 5
Private Const C_T0 As Long = 6                           ' six tier columns F..K
Private Const C_LAST As Long = 11
Private Const C_KEY As Long = 30                         ' hidden: the block's target on every row
Private Const PANEL_COL As Long = 8                      ' H at draw time
Private Const PANEL_W As Double = 560
Private Const PANEL_H As Double = 640

Private Const ZH_FONT As String = "Noto Sans TC"
Private Const MONO As String = "Consolas"
Private Const CLR_TEXT As Long = 14540253                ' 221,221,221
Private Const CLR_SOFT As Long = 12632256                ' 192,192,192
Private Const CLR_MUTED As Long = 8421504                ' 128,128,128
Private Const CLR_BANNER As Long = 1842204               ' 28,28,28
Private Const SORT_MARK As String = "THSORT"             ' "date" (default) / "ticker"

' ----------------------------------------------------------------
Private Function L(ByVal key As String) As String
    Select Case key
        Case "S1": L = ChrW(&H4E00) & ChrW(&H53E5) & ChrW(&H8A71) & ChrW(&H8AD6) & ChrW(&H9EDE)
        Case "S2": L = ChrW(&H5546) & ChrW(&H696D) & ChrW(&H6A21) & ChrW(&H5F0F) & ChrW(&H8207) & ChrW(&H8B77) & ChrW(&H57CE) & ChrW(&H6CB3)
        Case "S3": L = ChrW(&H70BA) & ChrW(&H4EC0) & ChrW(&H9EBC) & ChrW(&H662F) & ChrW(&H73FE) & ChrW(&H5728)
        Case "S4": L = ChrW(&H4E09) & ChrW(&H500B) & ChrW(&H652F) & ChrW(&H6490) & ChrW(&H6578) & ChrW(&H5B57)
        Case "S5": L = ChrW(&H6700) & ChrW(&H5927) & ChrW(&H98A8) & ChrW(&H96AA) & ChrW(&H8207) & ChrW(&H8B49) & ChrW(&H507D) & ChrW(&H8A0A) & ChrW(&H865F)
        Case "S6": L = ChrW(&H6C7A) & ChrW(&H7B56) & ChrW(&H8207) & ChrW(&H4E0B) & ChrW(&H4E00) & ChrW(&H500B) & ChrW(&H9A57) & ChrW(&H8B49) & ChrW(&H9EDE)
        Case "NEXT": L = ChrW(&H4E0B) & ChrW(&H6B21) & ChrW(&H9A57) & ChrW(&H8B49) & ChrW(&H65E5)
    End Select
End Function

Private Function StatusList() As String
    StatusList = "Robust,Solid,Growing,Slowing,Sluggish,Challenging,Contraction,Warning"
End Function

Private Function StatusColor(ByVal s As String) As Long
    Select Case LCase$(Trim$(s))
        Case "robust", "solid", "growing": StatusColor = RGB(80, 200, 120)
        Case "slowing", "sluggish", "challenging", "contraction", "warning": StatusColor = RGB(230, 90, 90)
        Case Else: StatusColor = CLR_SOFT
    End Select
End Function

Private Function RoleColor(ByVal s As String) As Long
    Select Case UCase$(Trim$(s))
        Case "MOAT": RoleColor = RR4_ACCENT
        Case "RISK": RoleColor = RGB(230, 90, 90)
        Case "CATALYST": RoleColor = RGB(0, 190, 240)
        Case Else: RoleColor = CLR_MUTED
    End Select
End Function

' Tier column headers (data sheet and view share them).  The D tier shows as a
' Greek capital delta on the view; the data-sheet header stays ASCII "D OWN".
Private Function TierName(ByVal t As Long, ByVal forView As Boolean) As String
    Select Case t
        Case 1: TierName = "T0 Q&A"
        Case 2: TierName = "T1 CALL"
        Case 3: TierName = "T1L FILING"
        Case 4: TierName = "T2 RESEARCH"
        Case 5: TierName = "T3 MEDIA"
        Case 6: TierName = IIf(forView, ChrW(&H394) & " OWN", "D OWN")
    End Select
End Function

Private Function TierColor(ByVal t As Long) As Long
    Select Case t
        Case 1: TierColor = RGB(245, 245, 245)           ' Q&A: the brightest
        Case 2, 3: TierColor = CLR_SOFT
        Case 4: TierColor = RGB(170, 170, 170)
        Case 5: TierColor = CLR_MUTED
        Case Else: TierColor = RGB(200, 170, 120)        ' own derivations: warm tint
    End Select
End Function

' ================================================================
'  TH! - migrate v1 once, make sure the data sheet exists, redraw
' ================================================================
Sub BuildThesisLibrary()
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fail
    Call MigrateV1ToArchive
    Call EnsureNotesSheet
    Dim ws As Worksheet: Set ws = EnsureViewSheet()
    Application.EnableEvents = prevEv
    Call DrawThesisView(ws)
    Call NavNotify("Thesis Library: " & NoteCount() & " notes")
    Exit Sub
Fail:
    Application.EnableEvents = prevEv
    Call NavNotify("Library failed: " & Err.Description, True)
End Sub

' v1 kept tblThesis on the "Thesis Library" sheet itself.  Rename that sheet to
' the archive (data untouched), hide it, drop its panel and v1 event code.
Private Sub MigrateV1ToArchive()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(THESIS_SHEET_OLD)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(THESIS_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If SheetExists(ARCHIVE_SHEET) Then Exit Sub
    On Error Resume Next
    ws.Shapes("TH_PANEL").Delete
    On Error GoTo 0
    ws.Name = ARCHIVE_SHEET
    Call WriteSheetCode(ws, "")                      ' no events on the archive
    ws.Visible = xlSheetHidden
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function EnsureViewSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(THESIS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then                              ' a v2 view still under the old tab name: just rename it
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(THESIS_SHEET_OLD)
        On Error GoTo 0
        If Not ws Is Nothing Then ws.Name = THESIS_SHEET
    End If
    If ws Is Nothing Then
        Dim prev As Object: Set prev = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = THESIS_SHEET
        ActiveWindow.DisplayGridlines = False
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If
    Call WriteSheetCode(ws, "Option Explicit" & vbCrLf & vbCrLf & _
        "' Thesis Library view (modThesis): QUERY / KEYWORD inputs, double-click header / target" & vbCrLf & _
        "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    Call ThesisViewChange(Me, Target)" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & vbCrLf & _
        "    Call ThesisViewDoubleClick(Me, Target, Cancel)" & vbCrLf & _
        "End Sub" & vbCrLf, "ThesisViewChange")
    Set EnsureViewSheet = ws
End Function

Private Sub EnsureNotesSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOTES_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Dim prev As Object: Set prev = ActiveSheet
        Dim after As Object
        On Error Resume Next
        Set after = ThisWorkbook.Worksheets(THESIS_SHEET)
        On Error GoTo 0
        If after Is Nothing Then Set after = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
        Set ws = ThisWorkbook.Worksheets.Add(After:=after)
        ws.Name = NOTES_SHEET
        ActiveWindow.DisplayGridlines = False
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If

    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = MONO: .Font.Size = 9: .Font.Color = CLR_TEXT
        .VerticalAlignment = xlCenter
    End With
    With ws.cells(1, 1)
        .Value = "THESIS NOTES  .  data for Library (L) - one theme note per row, its evidence in ONE tier column (T0 Q&A > T1 CALL > T1L FILING lag > T2 RESEARCH > T3 MEDIA > D OWN); type below the table to add, delete a row to remove"
        .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(NOTES_TABLE)
    On Error GoTo 0
    Dim hdr As Variant
    hdr = Array("TARGET", "TYPE", "CALL DATE", "STATUS", "ROLE", "THEME", "BEHAVIOR", _
                TierName(1, False), TierName(2, False), TierName(3, False), TierName(4, False), TierName(5, False), TierName(6, False))
    Dim j As Long
    If lo Is Nothing Then
        For j = 0 To NT_NCOL - 1: ws.cells(3, j + 1).Value = hdr(j): Next j
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.cells(3, 1), ws.cells(4, NT_NCOL)), , xlYes)
        lo.Name = NOTES_TABLE
    ElseIf lo.ListColumns.count = 8 Then
        ' 2026-09-15 upgrade: the old EVIDENCE column becomes T1L FILING (col 10);
        ' T0 / T1 go in front of it, T2 / T3 / D after.  Existing evidence stays
        ' where it is - re-sorting it into the right tier is a one-off outside job.
        lo.ListColumns.Add(NT_T0).Name = TierName(1, False)
        lo.ListColumns.Add(NT_T0 + 1).Name = TierName(2, False)
        lo.ListColumns(NT_T0 + 2).Name = TierName(3, False)
        lo.ListColumns.Add.Name = TierName(4, False)
        lo.ListColumns.Add.Name = TierName(5, False)
        lo.ListColumns.Add.Name = TierName(6, False)
    End If
    lo.TableStyle = ""
    With lo.HeaderRowRange
        .Font.Color = RR4_ACCENT: .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RR4_LINE
    End With
    Dim widths As Variant: widths = Array(16, 8, 12, 13, 10, 36, 56, 44, 44, 44, 44, 36, 44)
    For j = 0 To NT_NCOL - 1: ws.Columns(j + 1).ColumnWidth = widths(j): Next j
    If Not lo.DataBodyRange Is Nothing Then
        With lo.DataBodyRange
            .Interior.Color = RGB(8, 8, 8): .Font.Color = CLR_TEXT
        End With
        lo.ListColumns(NT_DATE).DataBodyRange.NumberFormat = "yyyy/mm/dd"
        lo.ListColumns(NT_TARGET).DataBodyRange.Font.Name = ZH_FONT
        For j = NT_THEME To NT_NCOL: lo.ListColumns(j).DataBodyRange.Font.Name = ZH_FONT: Next j
        Call SetList(lo.ListColumns(NT_TYPE).DataBodyRange, "stock,macro")
        Call SetList(lo.ListColumns(NT_STATUS).DataBodyRange, StatusList())
        Call SetList(lo.ListColumns(NT_ROLE).DataBodyRange, "MOAT,RISK,CATALYST")
    End If
    Call WriteSheetCode(ws, "Option Explicit" & vbCrLf & vbCrLf & _
        "' ThesisNotes data (modThesis): editing a note redraws the Thesis Library view" & vbCrLf & _
        "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    Call ThesisNotesChange(Me, Target)" & vbCrLf & _
        "End Sub" & vbCrLf, "ThesisNotesChange")
End Sub

Private Sub SetList(rng As Range, ByVal items As String)
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=items
    rng.Validation.ShowError = False
    On Error GoTo 0
End Sub

Private Function NotesTable() As ListObject
    On Error Resume Next
    Set NotesTable = ThisWorkbook.Worksheets(NOTES_SHEET).ListObjects(NOTES_TABLE)
End Function

Private Function NoteCount() As Long
    Dim lo As ListObject: Set lo = NotesTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim r As Range
    For Each r In lo.ListColumns(NT_TARGET).DataBodyRange.cells
        If Trim$(CStr(r.Value)) <> "" Then NoteCount = NoteCount + 1
    Next r
End Function

' ================================================================
'  Events
' ================================================================
Public Sub ThesisNotesChange(ws As Worksheet, ByVal Target As Range)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(NOTES_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.Range) Is Nothing Then Exit Sub
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    Dim c As Range, rr As Long, body As Range
    Set body = Intersect(Target, lo.DataBodyRange)
    If body Is Nothing Then Set body = lo.DataBodyRange.Rows(lo.ListRows.count)
    For Each c In body.Rows
        rr = c.Row - lo.HeaderRowRange.Row
        If rr >= 1 And rr <= lo.ListRows.count Then
            With lo.DataBodyRange
                If Trim$(CStr(.cells(rr, NT_TARGET).Value)) <> "" Then
                    If Trim$(CStr(.cells(rr, NT_TYPE).Value)) = "" Then .cells(rr, NT_TYPE).Value = "stock"
                    If Trim$(CStr(.cells(rr, NT_DATE).Value)) = "" Then .cells(rr, NT_DATE).Value = Date
                End If
            End With
        End If
    Next c
    Call SetList(lo.ListColumns(NT_TYPE).DataBodyRange, "stock,macro")
    Call SetList(lo.ListColumns(NT_STATUS).DataBodyRange, StatusList())
    Call SetList(lo.ListColumns(NT_ROLE).DataBodyRange, "MOAT,RISK,CATALYST")
    lo.ListColumns(NT_DATE).DataBodyRange.NumberFormat = "yyyy/mm/dd"
    On Error GoTo 0
    Application.EnableEvents = prevEv
    If SheetExists(THESIS_SHEET) Then Call DrawThesisView(ThisWorkbook.Worksheets(THESIS_SHEET))
End Sub

Public Sub ThesisViewChange(ws As Worksheet, ByVal Target As Range)
    Dim r As Long: r = PG_IN + NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim inputs As Range
    Set inputs = Union(ws.cells(r, C_TGT + lc), ws.cells(r, C_THEME + lc))
    If Intersect(Target, inputs) Is Nothing Then Exit Sub
    Call DrawThesisView(ws)
End Sub

Public Sub ThesisViewDoubleClick(ws As Worksheet, ByVal Target As Range, ByRef Cancel As Boolean)
    Dim t As Range: Set t = Target.cells(1, 1)
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    If t.Row = PG_HDR + off And t.Column = C_TGT + lc Then
        Cancel = True
        Dim cur As String: cur = SortMode(ws)
        Call SetSortMode(ws, IIf(cur = "ticker", "date", "ticker"))
        Call DrawThesisView(ws)
        Exit Sub
    End If
    If t.Row = PG_TITLE + off Then
        Cancel = True
        Call ShowArchive(ws, "")
        Exit Sub
    End If
    If t.Row >= PG_LIST + off And t.Column = C_TGT + lc Then
        Dim key As String: key = CStr(ws.cells(t.Row, C_KEY + lc).Value)
        If key <> "" Then
            Cancel = True
            Call ShowArchive(ws, key)
        End If
    End If
End Sub

Private Function SortMode(ws As Worksheet) As String
    SortMode = "date"
    Dim nm As Name
    For Each nm In ws.Names
        If Right$(nm.Name, Len(SORT_MARK) + 1) = "!" & SORT_MARK Then
            If InStr(nm.RefersTo, "ticker") > 0 Then SortMode = "ticker"
            Exit Function
        End If
    Next nm
End Function

Private Sub SetSortMode(ws As Worksheet, ByVal mode As String)
    ws.Names.Add Name:=SORT_MARK, RefersTo:="=""" & mode & """", Visible:=False
End Sub

' ================================================================
'  View
' ================================================================
Public Sub DrawThesisView(ws As Worksheet)
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fail

    ' keep the typed filters across the redraw
    Dim q As String, kw As String
    If NavHasRows(ws) Then
        q = Trim$(CStr(ws.cells(PG_IN + NavOffset(ws), C_TGT + NavLeft(ws)).Value))
        kw = Trim$(CStr(ws.cells(PG_IN + NavOffset(ws), C_THEME + NavLeft(ws)).Value))
    End If
    Dim mode As String: mode = SortMode(ws)

    Call NavStrip(ws)
    ws.cells.Clear
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = MONO: .Font.Size = 9: .Font.Color = CLR_TEXT
        .VerticalAlignment = xlCenter
        .RowHeight = 17
    End With
    Dim widths As Variant: widths = Array(18, 13, 10, 30, 52, 40, 40, 40, 40, 32, 40, 2)   ' BEHAVIOR + six tiers wrap
    Dim j As Long
    For j = 0 To 11: ws.Columns(j + 1).ColumnWidth = widths(j): Next j
    ws.Columns(C_KEY).Hidden = True

    ' ---- title + inputs (label above input) ----
    With ws.cells(PG_TITLE, 1)
        .Value = "LIBRARY"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 12
    End With
    ws.Rows(PG_TITLE).RowHeight = 22
    Call Lbl(ws.cells(PG_LBL, C_TGT), "QUERY")
    Call Lbl(ws.cells(PG_LBL, C_THEME), "KEYWORD")
    ws.Range(ws.cells(PG_IN, C_TGT), ws.cells(PG_IN, C_STATUS)).Merge
    ws.Range(ws.cells(PG_IN, C_THEME), ws.cells(PG_IN, C_BEHAV)).Merge
    Call InputBox_(ws.Range(ws.cells(PG_IN, C_TGT), ws.cells(PG_IN, C_STATUS)), q)
    Call InputBox_(ws.Range(ws.cells(PG_IN, C_THEME), ws.cells(PG_IN, C_BEHAV)), kw)
    ws.Rows(PG_IN).RowHeight = 20

    ' ---- data ----
    Dim lo As ListObject: Set lo = NotesTable()
    Dim n As Long
    Dim tg() As String, ty() As String, dt() As Date, st() As String, ro() As String, th() As String, be() As String, ev() As String
    ' ev(i, t) = note i's evidence in tier t (1..NTIER); at most one tier is filled
    n = ReadNotes(lo, tg, ty, dt, st, ro, th, be, ev)

    ' totals over everything
    Dim calls As Object: Set calls = CreateObject("Scripting.Dictionary")
    Dim tickers As Object: Set tickers = CreateObject("Scripting.Dictionary")
    Dim latest As Object: Set latest = CreateObject("Scripting.Dictionary")
    Dim kind As Object: Set kind = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As String
    For i = 1 To n
        k = UCase$(tg(i))
        calls(k & "|" & CLng(dt(i))) = True
        tickers(k) = True
        If Not latest.Exists(k) Then
            latest(k) = dt(i): kind(k) = ty(i)
        ElseIf dt(i) > latest(k) Then
            latest(k) = dt(i): kind(k) = ty(i)
        End If
    Next i
    With ws.cells(PG_TOTAL, 1)
        .Value = Format(n, "#,##0") & " notes . " & calls.count & " calls . " & tickers.count & " tickers" & _
                 "      (double-click a ticker = its archived thesis, the title = close)"
        .Font.Color = CLR_SOFT: .Font.Bold = True
    End With
    With ws.Range(ws.cells(PG_TOTAL, 1), ws.cells(PG_TOTAL, C_LAST)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE
    End With

    ' ---- filter: latest call per target; with a QUERY every call of the queried
    '      targets is drawn (newest first); KEYWORD narrows either way ----
    Dim qs As Variant, kws As Variant
    qs = SplitList(q): kws = SplitList(kw)
    Dim keep() As Boolean: ReDim keep(0 To n)
    Dim shown As Object: Set shown = CreateObject("Scripting.Dictionary")   ' target -> note count
    Dim shownCalls As Object: Set shownCalls = CreateObject("Scripting.Dictionary")
    Dim drawnNotes As Long
    For i = 1 To n
        k = UCase$(tg(i))
        If IsEmpty(qs) And IsEmpty(kws) Then
            keep(i) = False                                 ' nothing typed: draw no blocks (2026-09-15, user)
        ElseIf IsEmpty(qs) Then
            keep(i) = (dt(i) = latest(k))
        Else
            keep(i) = MatchesQuery(k, qs)
        End If
        If keep(i) And Not IsEmpty(kws) Then keep(i) = MatchesKeyword(th(i) & " " & be(i) & " " & AllTiers(ev, i), kws)
        If keep(i) Then
            shown(k) = shown(k) + 1
            shownCalls(k & "|" & CLng(dt(i))) = True
            drawnNotes = drawnNotes + 1
        End If
    Next i
    With ws.cells(PG_IN, C_T0)
        .Value = shownCalls.count & " calls drawn . " & drawnNotes & " notes"
        .Font.Color = RGB(80, 200, 120): .Font.Bold = True
    End With

    ' ---- header ----
    ws.cells(PG_HDR, C_TGT).Value = IIf(mode = "ticker", "TICKER A-Z", "LATEST CALL")
    ws.cells(PG_HDR, C_STATUS).Value = "STATUS"
    ws.cells(PG_HDR, C_ROLE).Value = "ROLE"
    ws.cells(PG_HDR, C_THEME).Value = "THEME"
    ws.cells(PG_HDR, C_BEHAV).Value = "BEHAVIOR"
    Dim t As Long
    For t = 1 To NTIER: ws.cells(PG_HDR, C_T0 + t - 1).Value = TierName(t, True): Next t
    With ws.Range(ws.cells(PG_HDR, 1), ws.cells(PG_HDR, C_LAST))
        .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With
    ws.cells(PG_HDR, C_T0 + 2).AddComment "Lag information: statutory filings / official data - true but already old when published"
    ws.cells(PG_HDR, C_LAST).AddComment "Own derivations: cross-check tables, valuation page, report tabs, framework claims"
    ws.cells(PG_HDR, C_TGT).AddComment "Double-click: sort blocks by latest call date <-> ticker A-Z"

    ' ---- blocks ----
    Dim r As Long: r = PG_LIST
    Dim sec As Variant
    For Each sec In Array("stock", "macro")
        Dim order As Variant: order = OrderedTargets(shown, latest, kind, CStr(sec), mode)
        If Not IsEmpty(order) Then
            With ws.Range(ws.cells(r, 1), ws.cells(r, C_LAST))
                .Interior.Color = CLR_BANNER
                .Font.Color = RR4_ACCENT: .Font.Bold = True
            End With
            ws.cells(r, 1).Value = UCase$(CStr(sec))
            r = r + 2
            Dim ti As Long
            For ti = LBound(order) To UBound(order)
                r = DrawBlock(ws, r, CStr(order(ti)), latest(order(ti)), n, tg, dt, st, ro, th, be, ev, keep) + 1
            Next ti
        End If
    Next sec
    If r > PG_LIST Then                                     ' long BEHAVIOR / EVIDENCE wrap: let the rows grow
        ws.Range(ws.cells(PG_LIST, C_BEHAV), ws.cells(r, C_LAST)).WrapText = True
        ws.Range(ws.cells(PG_LIST, 1), ws.cells(r, C_LAST)).VerticalAlignment = xlTop
        ws.Range(ws.Rows(PG_LIST), ws.Rows(r)).AutoFit
        Dim wrapRow As Long
        For wrapRow = PG_LIST To r
            If ws.Rows(wrapRow).RowHeight < 17 Then ws.Rows(wrapRow).RowHeight = 17
        Next wrapRow
    End If
    If shown.count = 0 Then
        If n = 0 Then
            ws.cells(r, 1).Value = "No notes yet - add them on the ThesisNotes sheet."
        ElseIf IsEmpty(qs) And IsEmpty(kws) Then
            ws.cells(r, 1).Value = "Type a ticker in QUERY (comma = several) or a word in KEYWORD to show notes."
        Else
            ws.cells(r, 1).Value = "Nothing matches QUERY / KEYWORD."
        End If
        ws.cells(r, 1).Font.Color = CLR_MUTED
    End If

    Call EnsurePanel(ws)
    Call NavAdd(ws, "L")
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
    Exit Sub
Fail:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    Call NavAdd(ws, "L")
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
    Call NavNotify("Thesis Library draw failed: " & msg, True)
End Sub

Private Function DrawBlock(ws As Worksheet, ByVal r As Long, ByVal key As String, ByVal callDate As Date, ByVal n As Long, _
        tg() As String, dt() As Date, st() As String, ro() As String, th() As String, be() As String, ev() As String, keep() As Boolean) As Long
    Dim top As Long: top = r
    Dim i As Long, shownName As String
    ' every kept note of this target, newest call first (stable: sheet order within one date)
    Dim idx() As Long, m As Long, a As Long, b As Long, t As Long
    ReDim idx(1 To n + 1)
    For i = 1 To n
        If keep(i) And UCase$(tg(i)) = key Then
            m = m + 1: idx(m) = i
            b = m
            Do While b > 1
                If dt(idx(b - 1)) >= dt(idx(b)) Then Exit Do
                t = idx(b - 1): idx(b - 1) = idx(b): idx(b) = t
                b = b - 1
            Loop
        End If
    Next i
    Dim prevDt As Date, firstGroup As Boolean: firstGroup = True
    For a = 1 To m
        i = idx(a)
        If shownName = "" Then shownName = tg(i)
        If a > 1 And dt(i) <> prevDt Then                  ' a new (older) call starts: label it on its first row
            If firstGroup And r = top + 1 Then             ' first call had a single note: keep a row for its date
                ws.cells(r, C_KEY).Value = key
                r = r + 1
            End If
            firstGroup = False
            Call DateCell(ws.cells(r, C_TGT), dt(i))
        End If
        prevDt = dt(i)
        If True Then
            Dim sc As Range: Set sc = ws.cells(r, C_STATUS)
            sc.Value = Dash(st(i)): sc.Font.Color = StatusColor(st(i))
            With ws.cells(r, C_ROLE)
                .Value = Dash(ro(i)): .Font.Color = RoleColor(ro(i)): .HorizontalAlignment = xlCenter
            End With
            Call TextCell(ws.cells(r, C_THEME), th(i), RGB(245, 245, 245))
            Call TextCell(ws.cells(r, C_BEHAV), be(i), CLR_SOFT)
            Dim tt As Long                                 ' evidence in its tier column, the others stay blank
            For tt = 1 To NTIER
                If Trim$(ev(i, tt)) <> "" Then Call TextCell(ws.cells(r, C_T0 + tt - 1), ev(i, tt), TierColor(tt))
            Next tt
            ws.cells(r, C_KEY).Value = key
            r = r + 1
        End If
    Next a
    If r = top + 1 Then                                    ' room for the date line
        ws.cells(r, C_KEY).Value = key
        r = r + 1
    End If
    With ws.cells(top, C_TGT)
        .NumberFormat = "@"                              ' TW codes are digits: keep them text, left-aligned
        .HorizontalAlignment = xlLeft
        .Value = shownName
        On Error Resume Next
        .Errors(xlNumberAsText).Ignore = True            ' no green "number stored as text" flag
        On Error GoTo 0
        .Font.Name = ZH_FONT: .Font.Bold = True: .Font.Size = 10: .Font.Color = RGB(245, 245, 245)
    End With
    Call DateCell(ws.cells(top + 1, C_TGT), callDate)
    DrawBlock = r
End Function

Private Sub DateCell(cell As Range, ByVal d As Date)
    With cell
        .NumberFormat = "@"                              ' else Excel turns the text back into a date
        .Value = Format$(d, "yyyy-mm-dd")
        .Font.Color = CLR_MUTED
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub TextCell(cell As Range, ByVal v As String, ByVal clr As Long)
    If Trim$(v) = "" Then
        cell.Value = "-": cell.Font.Color = CLR_MUTED
    Else
        cell.Value = v: cell.Font.Color = clr
        cell.Font.Name = ZH_FONT
    End If
End Sub

Private Function Dash(ByVal v As String) As String
    If Trim$(v) = "" Then Dash = "-" Else Dash = v
End Function

Private Sub Lbl(cell As Range, ByVal txt As String)
    With cell
        .Value = txt
        .Font.Color = RR4_ACCENT: .Font.Bold = True
        .VerticalAlignment = xlBottom
    End With
End Sub

Private Sub InputBox_(rng As Range, ByVal v As String)
    With rng
        .NumberFormat = "@"
        .Value = v
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .Font.Name = ZH_FONT
        .HorizontalAlignment = xlLeft
    End With
End Sub

' Reads tblNotes into parallel 1-based arrays; rows without a TARGET are skipped.
Private Function ReadNotes(lo As ListObject, tg() As String, ty() As String, dt() As Date, st() As String, _
        ro() As String, th() As String, be() As String, ev() As String) As Long
    ReDim tg(0 To 0): ReDim ty(0 To 0): ReDim dt(0 To 0): ReDim st(0 To 0)
    ReDim ro(0 To 0): ReDim th(0 To 0): ReDim be(0 To 0): ReDim ev(0 To 0, 1 To NTIER)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = lo.DataBodyRange.Value
    Dim rows As Long: rows = UBound(v, 1)
    ReDim tg(0 To rows): ReDim ty(0 To rows): ReDim dt(0 To rows): ReDim st(0 To rows)
    ReDim ro(0 To rows): ReDim th(0 To rows): ReDim be(0 To rows): ReDim ev(0 To rows, 1 To NTIER)
    Dim ncol As Long: ncol = UBound(v, 2)
    Dim i As Long, n As Long
    For i = 1 To rows
        If Trim$(CStr(v(i, NT_TARGET))) <> "" Then
            n = n + 1
            tg(n) = Trim$(CStr(v(i, NT_TARGET)))
            ty(n) = LCase$(Trim$(CStr(v(i, NT_TYPE))))
            If ty(n) <> "macro" Then ty(n) = "stock"
            If IsDate(v(i, NT_DATE)) Then dt(n) = CDate(v(i, NT_DATE)) Else dt(n) = 0
            st(n) = Trim$(CStr(v(i, NT_STATUS)))
            ro(n) = UCase$(Trim$(CStr(v(i, NT_ROLE))))
            th(n) = CStr(v(i, NT_THEME))
            be(n) = CStr(v(i, NT_BEHAV))
            Dim t As Long
            For t = 1 To NTIER
                If NT_T0 + t - 1 <= ncol Then ev(n, t) = CStr(v(i, NT_T0 + t - 1))
            Next t
        End If
    Next i
    ReadNotes = n
End Function

' All tiers of one note joined, for KEYWORD matching.
Private Function AllTiers(ev() As String, ByVal i As Long) As String
    Dim t As Long
    For t = 1 To NTIER: AllTiers = AllTiers & " " & ev(i, t): Next t
End Function

Private Function SplitList(ByVal s As String) As Variant
    SplitList = Empty
    If Trim$(s) = "" Then Exit Function
    Dim parts As Variant: parts = Split(s, ",")
    Dim out() As String, i As Long, n As Long
    ReDim out(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Trim$(parts(i)) <> "" Then out(n) = UCase$(Trim$(parts(i))): n = n + 1
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve out(0 To n - 1)
    SplitList = out
End Function

Private Function MatchesQuery(ByVal key As String, qs As Variant) As Boolean
    Dim q As Variant, bare As String
    bare = key
    If Right$(bare, 4) = ".TWO" Then bare = Left$(bare, Len(bare) - 4)
    If Right$(bare, 3) = ".TW" Then bare = Left$(bare, Len(bare) - 3)
    For Each q In qs
        If key = q Or bare = q Then MatchesQuery = True: Exit Function
    Next q
End Function

Private Function MatchesKeyword(ByVal text As String, kws As Variant) As Boolean
    Dim w As Variant, t As String: t = UCase$(text)
    For Each w In kws
        If InStr(t, CStr(w)) > 0 Then MatchesKeyword = True: Exit Function
    Next w
End Function

' Targets of one section, ordered by latest call (newest first) or A-Z.
Private Function OrderedTargets(shown As Object, latest As Object, kind As Object, ByVal sec As String, ByVal mode As String) As Variant
    OrderedTargets = Empty
    Dim ks() As String, n As Long, k As Variant
    ReDim ks(1 To shown.count + 1)
    For Each k In shown.keys
        If kind(k) = sec Then n = n + 1: ks(n) = CStr(k)
    Next k
    If n = 0 Then Exit Function
    Dim i As Long, j As Long, tmp As String, swap As Boolean
    For i = 1 To n - 1
        For j = i + 1 To n
            If mode = "ticker" Then
                swap = (ks(j) < ks(i))
            Else
                swap = (latest(ks(j)) > latest(ks(i)))
                If latest(ks(j)) = latest(ks(i)) Then swap = (ks(j) < ks(i))
            End If
            If swap Then tmp = ks(i): ks(i) = ks(j): ks(j) = tmp
        Next j
    Next i
    ReDim Preserve ks(1 To n)
    OrderedTargets = ks
End Function

' ================================================================
'  Reading panel: the archived six-section thesis of a target
' ================================================================
Private Sub EnsurePanel(ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("TH_PANEL")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, ws.Columns(C_BEHAV).Left, ws.Rows(PG_HDR).Top, PANEL_W, PANEL_H)
        shp.Name = "TH_PANEL"
    End If
    shp.Width = PANEL_W
    shp.Placement = xlMove
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(12, 12, 12)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RR4_LINE
    shp.Line.Weight = 0.75
    With shp.TextFrame2
        .MarginLeft = 12: .MarginRight = 12: .MarginTop = 10: .MarginBottom = 10
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeShapeToFitText
        .VerticalAnchor = msoAnchorTop
    End With
    shp.Visible = msoFalse                           ' shown on demand by ShowArchive
End Sub

Private Sub ShowArchive(ws As Worksheet, ByVal key As String)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("TH_PANEL")
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    If key = "" Then shp.Visible = msoFalse: Exit Sub
    ' float over BEHAVIOR / EVIDENCE, from the header row down (a reading overlay)
    shp.Left = ws.Columns(C_BEHAV + NavLeft(ws)).Left
    shp.Top = ws.Rows(PG_HDR + NavOffset(ws)).Top
    shp.Visible = msoTrue
    Dim tr As Object: Set tr = shp.TextFrame2.TextRange

    Dim lo As ListObject, row As Long
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(ARCHIVE_SHEET).ListObjects(THESIS_TABLE)
    On Error GoTo 0
    If Not lo Is Nothing Then row = ArchiveRow(lo, key)

    If row = 0 Then
        tr.Text = key & vbCr & "No archived thesis for this target - its notes are all there is." & vbCr & _
                  "(double-click the page title to close)"
        tr.Font.Name = MONO: tr.Font.NameFarEast = ZH_FONT: tr.Font.Size = 10
        tr.Font.Bold = msoFalse: tr.Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
        Exit Sub
    End If

    Dim b As Range: Set b = lo.DataBodyRange
    Dim secKeys As Variant: secKeys = Array("S1", "S2", "S3", "S4", "S5", "S6")
    Dim txt As String, starts(0 To 7) As Long, lens(0 To 7) As Long
    Dim head As String: head = CStr(b.cells(row, 3).Value) & "   " & CStr(b.cells(row, 4).Value)
    starts(0) = 1: lens(0) = Len(head)
    txt = head & vbCr
    Dim meta As String
    meta = CStr(b.cells(row, 2).Value) & "  .  " & CStr(b.cells(row, 5).Value) & "  .  " & CStr(b.cells(row, 13).Value) & _
           "  .  " & Format(b.cells(row, 1).Value, "yyyy/mm/dd")
    If CStr(b.cells(row, 12).Value) <> "" Then meta = meta & "  .  " & L("NEXT") & " " & Format(b.cells(row, 12).Value, "yyyy/mm/dd")
    starts(1) = Len(txt) + 1: lens(1) = Len(meta)
    txt = txt & meta & vbCr & vbCr
    Dim k As Long
    For k = 0 To 5
        Dim h As String: h = (k + 1) & "  " & L(CStr(secKeys(k)))
        starts(k + 2) = Len(txt) + 1: lens(k + 2) = Len(h)
        Dim body As String: body = Trim$(CStr(b.cells(row, 6 + k).Value))
        If body = "" Then body = "-"
        txt = txt & h & vbCr & body & vbCr & vbCr
    Next k
    tr.Text = txt
    tr.Font.Name = ZH_FONT: tr.Font.NameFarEast = ZH_FONT
    tr.Font.Size = 10: tr.Font.Bold = msoFalse
    tr.Font.Fill.ForeColor.RGB = RGB(215, 215, 215)
    tr.ParagraphFormat.SpaceAfter = 2
    With tr.Characters(starts(0), lens(0)).Font
        .Size = 13: .Bold = msoTrue: .Fill.ForeColor.RGB = RR4_ACCENT
    End With
    With tr.Characters(starts(1), lens(1)).Font
        .Size = 9: .Fill.ForeColor.RGB = RGB(140, 140, 140)
    End With
    For k = 2 To 7
        With tr.Characters(starts(k), lens(k)).Font
            .Bold = msoTrue: .Fill.ForeColor.RGB = RR4_ACCENT: .Size = 10
        End With
    Next k
End Sub

' Latest archived thesis row for key (ties: the lower row, i.e. added later).
Private Function ArchiveRow(lo As ListObject, ByVal key As String) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = lo.DataBodyRange.Value
    Dim i As Long, bestD As Double, d As Double
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(CStr(v(i, 3)))) = key Then
            d = 0
            If IsDate(v(i, 1)) Then d = CDbl(CDate(v(i, 1)))
            If ArchiveRow = 0 Or d >= bestD Then ArchiveRow = i: bestD = d
        End If
    Next i
End Function

' ================================================================
' Sheet event code, written into the document module (needs "Trust access to
' the VBA project object model").  code = "" empties the module.  marker: skip
' when the module already contains it.
Private Sub WriteSheetCode(ws As Worksheet, ByVal code As String, Optional ByVal marker As String = "")
    On Error GoTo Skip
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 Then
            If comp.Properties("Name").Value = ws.Name Then
                Dim cm As Object: Set cm = comp.CodeModule
                If cm.CountOfLines > 0 Then
                    If marker <> "" Then
                        If InStr(cm.Lines(1, cm.CountOfLines), marker) > 0 Then Exit Sub
                    End If
                    cm.DeleteLines 1, cm.CountOfLines
                End If
                If code <> "" Then cm.AddFromString code
                Exit Sub
            End If
        End If
    Next comp
    Exit Sub
Skip:
    Call NavNotify("Thesis Library: sheet event code for '" & ws.Name & "' could not be written (Trust access to the VBA project object model)", True)
End Sub
