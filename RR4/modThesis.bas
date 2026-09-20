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
'      T3 MEDIA | D OWN | SOURCE | SUBTYPE
'      (2026-09-19: SUBTYPE industry / macro splits TYPE = macro notes; the view
'      draws STOCK, INDUSTRY, MACRO sections.  A macro note typed without a
'      SUBTYPE prompts Yes = INDUSTRY / No = MACRO / Cancel = blank.)
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
'      (2026-09-19: 14th column SOURCE added - a citation for wherever the
'      note's content came from.  Most rows hold a local file path (a PDF
'      under the user's Bloomberg folder, or a path into this repo's
'      research/ reports); when there is no saved file (e.g. a note built
'      from a pasted newsletter/Substack article) a free-text citation goes
'      there instead - publication, date, title.  EnsureNotesSheet upgrades
'      a 13-column table in place by appending SOURCE, same pattern as the
'      2026-09-15 tier split; an 8-column table upgrades straight to 14 in
'      one pass.  This column was added directly on the live workbook before
'      modThesis.bas caught up - see CLAUDE.md's RR4 section for the story.)
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
'                  latest call date (newest first) <-> ticker A-Z (THSORT).
'                  Double-click STATUS or ROLE = sort the notes INSIDE every
'                  block by that column (2026-09-15): first click list order
'                  (STATUS green Robust..Growing then red Slowing..Warning;
'                  ROLE MOAT > RISK > CATALYST), second click reversed, third
'                  click off; blanks always last, ties keep data-sheet order.
'                  While a note sort is on the call-date grouping of a block
'                  is dropped (older calls mix in, no date labels).  Kept in
'                  the hidden name THNOTESORT ("status|asc" etc.), header
'                  shows a down / up triangle.
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
'  2026-09-16 (workbench pass, three user-picked items):
'    1. Rows are a fixed three lines high (NOTE_ROW_H); long BEHAVIOR /
'       evidence is clipped, not auto-fitted.  Double-click any cell of a
'       note row -> the reading panel (same TH_PANEL, same spot) shows that
'       note in full: header, THEME, BEHAVIOR, its tier + evidence, then the
'       target's archived thesis one-liner and next-check section as a
'       reminder.  Double-click the target cell still shows the whole archive.
'       Hidden column C_ROW carries the note's tblNotes body row.
'    2. Empty state (nothing typed) draws an INDEX instead of a hint: every
'       target in four column groups (target . notes . latest call), STOCK
'       then MACRO, newest call first (THSORT "ticker" = A-Z).  The date is
'       greyed after FRESH_GREY days and orange after FRESH_ORANGE days.
'       Double-click a target in the index = put it in QUERY and draw it.
'    3. Third input FILTER (F3:G3, space or comma separated, tokens are OR-ed,
'       i.e. a note passes if ANY token matches): STATUS words, GREEN / RED,
'       ROLE words, tiers T0 T1 T1L T2 T3 D, >yyyy-mm-dd / <yyyy-mm-dd,
'       or a yyyy-mm / yyyy prefix of the call date.  It narrows whatever
'       QUERY / KEYWORD produced (AND against them).  QUERY also takes
'       @PORT (tickers in the RR4 position log) and @WATCH (the WATCHLIST).
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
Private Const NT_SOURCE As Long = 14                      ' citation: local file path, or free text when there is none
Private Const NT_SUBTYPE As Long = 15                     ' 2026-09-19: macro notes split into industry / macro
Private Const NT_NCOL As Long = 15

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
' 2026-09-20: the six tier columns (T0..D) that used to sit side by side after
' BEHAVIOR are gone from the grid - each note's evidence still lives in exactly
' one of tblNotes' six tier columns (NT_T0..NT_T0+NTIER-1, unchanged), but the
' view now only shows which one via a single narrow TIER badge (see TierShort),
' and the full evidence text only surfaces in the reading panel (ShowNote).
' BEHAVIOR takes the width that freed up.
Private Const C_TIER As Long = 4                         ' badge: which tier this note's evidence sits in
Private Const C_THEME As Long = 5
Private Const C_BEHAV As Long = 6
Private Const C_LAST As Long = 6
Private Const C_FILTER As Long = 4                       ' FILTER input reuses TIER/THEME's row-3 slot, merged D3:E3; "N calls drawn" moved to F3 (BEHAVIOR's slot)
Private Const C_COUNT As Long = 6
Private Const C_KEY As Long = 30                         ' hidden: the block's target on every row
Private Const C_ROW As Long = 31                         ' hidden: the note's tblNotes body row (note rows only)
Private Const NOTE_ROW_H As Double = 20                  ' 2026-09-20: was 42 (3 lines); most notes only used 1-2, clipped beyond that
Private Const FRESH_GREY As Long = 60                    ' index: latest call older than this = grey
Private Const FRESH_ORANGE As Long = 120                 ' ... older than this = orange
Private Const IDX_GROUPS As Long = 4                     ' index column groups (target . notes . latest) across the page
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
Private Const NOTESORT_MARK As String = "THNOTESORT"     ' "" / "status|asc" / "status|desc" / "role|asc" / "role|desc"

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

' Short code for the TIER badge column (2026-09-20) - just enough to tell tiers
' apart at a glance; the full name only shows in the reading panel.
Private Function TierShort(ByVal t As Long) As String
    Select Case t
        Case 1: TierShort = "T0"
        Case 2: TierShort = "T1"
        Case 3: TierShort = "T1L"
        Case 4: TierShort = "T2"
        Case 5: TierShort = "T3"
        Case 6: TierShort = ChrW(&H394)
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
                TierName(1, False), TierName(2, False), TierName(3, False), TierName(4, False), TierName(5, False), TierName(6, False), _
                "SOURCE", "SUBTYPE")
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
        lo.ListColumns.Add.Name = "SOURCE"
    ElseIf lo.ListColumns.count = 13 Then
        ' 2026-09-19 upgrade: append SOURCE (citation - local file path, or free
        ' text when there is none), same append-only pattern as the tier split above.
        lo.ListColumns.Add.Name = "SOURCE"
    End If
    If lo.ListColumns.count = 14 Then
        ' 2026-09-19 upgrade: append SUBTYPE (industry / macro) for TYPE = macro notes
        lo.ListColumns.Add.Name = "SUBTYPE"
    End If
    lo.TableStyle = ""
    With lo.HeaderRowRange
        .Font.Color = RR4_ACCENT: .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RR4_LINE
    End With
    Dim widths As Variant: widths = Array(16, 8, 12, 13, 10, 36, 56, 44, 44, 44, 44, 36, 44, 50, 10)
    For j = 0 To NT_NCOL - 1: ws.Columns(j + 1).ColumnWidth = widths(j): Next j
    If Not lo.DataBodyRange Is Nothing Then
        With lo.DataBodyRange
            .Interior.Color = RGB(8, 8, 8): .Font.Color = CLR_TEXT
        End With
        lo.ListColumns(NT_DATE).DataBodyRange.NumberFormat = "yyyy/mm/dd"
        lo.ListColumns(NT_TARGET).DataBodyRange.Font.Name = ZH_FONT
        For j = NT_THEME To NT_NCOL: lo.ListColumns(j).DataBodyRange.Font.Name = ZH_FONT: Next j
        Call SetList(lo.ListColumns(NT_TYPE).DataBodyRange, "stock,macro")
        Call SetList(lo.ListColumns(NT_SUBTYPE).DataBodyRange, "industry,macro")
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
                    ' 2026-09-19: user wants to be asked every time a macro note has no SUBTYPE
                    If LCase$(Trim$(CStr(.cells(rr, NT_TYPE).Value))) = "macro" _
                       And Trim$(CStr(.cells(rr, NT_SUBTYPE).Value)) = "" Then
                        Call AskSubtype(lo, rr)
                    End If
                End If
            End With
        End If
    Next c
    Call SetList(lo.ListColumns(NT_TYPE).DataBodyRange, "stock,macro")
    Call SetList(lo.ListColumns(NT_SUBTYPE).DataBodyRange, "industry,macro")
    Call SetList(lo.ListColumns(NT_STATUS).DataBodyRange, StatusList())
    Call SetList(lo.ListColumns(NT_ROLE).DataBodyRange, "MOAT,RISK,CATALYST")
    lo.ListColumns(NT_DATE).DataBodyRange.NumberFormat = "yyyy/mm/dd"
    On Error GoTo 0
    Application.EnableEvents = prevEv
    If SheetExists(THESIS_SHEET) Then Call DrawThesisView(ThisWorkbook.Worksheets(THESIS_SHEET))
End Sub

' Asks whether macro note rr (tblNotes body row) is INDUSTRY or MACRO; the
' hint is what other notes of the same target already use.  Cancel leaves it
' blank (drawn under the target's existing section, else MACRO).
Private Sub AskSubtype(lo As ListObject, ByVal rr As Long)
    Dim tgt As String: tgt = UCase$(Trim$(CStr(lo.DataBodyRange.cells(rr, NT_TARGET).Value)))
    Dim v As Variant: v = lo.DataBodyRange.Value
    Dim i As Long, hint As String, s As String
    For i = 1 To UBound(v, 1)
        If i <> rr And UCase$(Trim$(CStr(v(i, NT_TARGET)))) = tgt Then
            s = LCase$(Trim$(CStr(v(i, NT_SUBTYPE))))
            If s = "industry" Or s = "macro" Then hint = UCase$(s): Exit For
        End If
    Next i
    Dim msg As String
    msg = "Macro note '" & tgt & "' has no SUBTYPE." & vbLf & vbLf & _
          "Yes = INDUSTRY    No = MACRO    Cancel = leave blank"
    If hint <> "" Then msg = msg & vbLf & vbLf & "Other notes of this target: " & hint
    Select Case MsgBox(msg, vbYesNoCancel + vbQuestion, "Library SUBTYPE")
        Case vbYes: lo.DataBodyRange.cells(rr, NT_SUBTYPE).Value = "industry"
        Case vbNo: lo.DataBodyRange.cells(rr, NT_SUBTYPE).Value = "macro"
    End Select
End Sub

Public Sub ThesisViewChange(ws As Worksheet, ByVal Target As Range)
    Dim r As Long: r = PG_IN + NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim inputs As Range
    Set inputs = Union(ws.cells(r, C_TGT + lc), ws.cells(r, C_THEME + lc), ws.cells(r, C_FILTER + lc))
    If Intersect(Target, inputs) Is Nothing Then Exit Sub
    Call DrawThesisView(ws)
End Sub

' True while the page shows the index (all three inputs blank).
Private Function InIndexMode(ws As Worksheet) As Boolean
    If Not NavHasRows(ws) Then Exit Function
    Dim r As Long: r = PG_IN + NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    InIndexMode = (Trim$(CStr(ws.cells(r, C_TGT + lc).Value)) = "" And Trim$(CStr(ws.cells(r, C_THEME + lc).Value)) = "" _
                   And Trim$(CStr(ws.cells(r, C_FILTER + lc).Value)) = "")
End Function

Public Sub ThesisViewDoubleClick(ws As Worksheet, ByVal Target As Range, ByRef Cancel As Boolean)
    Dim t As Range: Set t = Target.cells(1, 1)
    Dim off As Long: off = NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    If t.Row = PG_HDR + off And t.Column = C_TGT + lc Then
        Cancel = True
        If Trim$(CStr(t.Value)) = "" Then Exit Sub          ' index view has no header row (2026-09-19)
        Dim cur As String: cur = SortMode(ws)
        Call SetSortMode(ws, IIf(cur = "ticker", "date", "ticker"))
        Call DrawThesisView(ws)
        Exit Sub
    End If
    If t.Row = PG_HDR + off And (t.Column = C_STATUS + lc Or t.Column = C_ROLE + lc) Then
        Cancel = True
        Dim nk As String: nk = IIf(t.Column = C_STATUS + lc, "status", "role")
        Dim ns As String: ns = NoteSort(ws)
        If ns = nk & "|asc" Then
            ns = nk & "|desc"
        ElseIf ns = nk & "|desc" Then
            ns = ""
        Else
            ns = nk & "|asc"
        End If
        Call SetNoteSort(ws, ns)
        Call DrawThesisView(ws)
        Exit Sub
    End If
    If t.Row = PG_TITLE + off Then
        Cancel = True
        Call ShowArchive(ws, "", 0)
        Exit Sub
    End If
    If t.Row < PG_LIST + off Then Exit Sub
    ' index: double-click a target = query it
    If InIndexMode(ws) Then
        Dim g As Long
        For g = 0 To IDX_GROUPS - 1
            If t.Column = 1 + g * 3 + lc Then
                Dim tk As String: tk = Trim$(CStr(t.Value))
                If tk <> "" And ws.cells(t.Row, C_KEY + lc).Value <> "" Then
                    Cancel = True
                    Dim prevEv As Boolean: prevEv = Application.EnableEvents
                    Application.EnableEvents = False
                    ws.cells(PG_IN + off, C_TGT + lc).Value = tk
                    Application.EnableEvents = prevEv
                    Call DrawThesisView(ws)
                End If
                Exit Sub
            End If
        Next g
        Exit Sub
    End If
    Dim key As String: key = CStr(ws.cells(t.Row, C_KEY + lc).Value)
    If key = "" Then Exit Sub
    If t.Column = C_TGT + lc Then
        Cancel = True
        Call ShowArchive(ws, key, t.Row)
    ElseIf t.Column > C_TGT + lc And t.Column <= C_LAST + lc Then
        Dim body As String: body = CStr(ws.cells(t.Row, C_ROW + lc).Value)
        If body <> "" Then
            Cancel = True
            Call ShowNote(ws, key, CLng(body), t.Row)
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

' Note sort inside the blocks: "" (data-sheet order) or "status|asc" etc.
Private Function NoteSort(ws As Worksheet) As String
    Dim nm As Name
    For Each nm In ws.Names
        If Right$(nm.Name, Len(NOTESORT_MARK) + 1) = "!" & NOTESORT_MARK Then
            NoteSort = Replace(Replace(nm.RefersTo, "=", ""), """", "")
            Exit Function
        End If
    Next nm
End Function

Private Sub SetNoteSort(ws As Worksheet, ByVal mode As String)
    ws.Names.Add Name:=NOTESORT_MARK, RefersTo:="=""" & mode & """", Visible:=False
End Sub

' Rank of one note under the note sort (lower = higher on the page).  Blank /
' unknown values rank last whichever direction; the reverse direction flips
' only the known values.
Private Function NoteRank(ByVal mode As String, ByVal status As String, ByVal role As String) As Long
    Dim r As Long, nKnown As Long
    If Left$(mode, 6) = "status" Then
        nKnown = 8
        Select Case LCase$(Trim$(status))
            Case "robust": r = 1
            Case "solid": r = 2
            Case "growing": r = 3
            Case "slowing": r = 4
            Case "sluggish": r = 5
            Case "challenging": r = 6
            Case "contraction": r = 7
            Case "warning": r = 8
        End Select
    Else
        nKnown = 3
        Select Case UCase$(Trim$(role))
            Case "MOAT": r = 1
            Case "RISK": r = 2
            Case "CATALYST": r = 3
        End Select
    End If
    If r = 0 Then
        NoteRank = 99
    ElseIf Right$(mode, 4) = "desc" Then
        NoteRank = nKnown + 1 - r
    Else
        NoteRank = r
    End If
End Function

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
    Dim q As String, kw As String, fl As String
    If NavHasRows(ws) Then
        q = Trim$(CStr(ws.cells(PG_IN + NavOffset(ws), C_TGT + NavLeft(ws)).Value))
        kw = Trim$(CStr(ws.cells(PG_IN + NavOffset(ws), C_THEME + NavLeft(ws)).Value))
        fl = Trim$(CStr(ws.cells(PG_IN + NavOffset(ws), C_FILTER + NavLeft(ws)).Value))
        If fl Like "* calls drawn *" Then fl = ""            ' pre-2026-09-16 pages had the count text in this cell
    End If
    Dim mode As String: mode = SortMode(ws)
    Dim indexMode As Boolean: indexMode = (q = "" And kw = "" And fl = "")

    Call NavStrip(ws)
    ws.cells.Clear
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = MONO: .Font.Size = 9: .Font.Color = CLR_TEXT
        .VerticalAlignment = xlCenter
        .RowHeight = 17
    End With
    ' 2026-09-20: TIER is a narrow badge (see TierShort); the six tier columns it
    ' replaced free their width to BEHAVIOR. Columns 7-12 are unused spacers now,
    ' kept at width 2 so the array stays 12 long (indexMode's own 12-wide layout
    ' below reuses the same "For j = 0 To 11" loop).
    Dim widths As Variant: widths = Array(18, 13, 10, 6, 30, 220, 2, 2, 2, 2, 2, 2)
    If indexMode Then widths = Array(18, 8, 12, 18, 8, 12, 18, 8, 12, 18, 8, 12)     ' four (target . notes . latest) groups
    Dim j As Long
    For j = 0 To 11: ws.Columns(j + 1).ColumnWidth = widths(j): Next j
    ws.Columns(C_KEY).Hidden = True
    ws.Columns(C_ROW).Hidden = True

    ' ---- title + inputs (label above input) ----
    With ws.cells(PG_TITLE, 1)
        .Value = "LIBRARY"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 12
    End With
    ws.Rows(PG_TITLE).RowHeight = 22
    Call Lbl(ws.cells(PG_LBL, C_TGT), "QUERY")
    Call Lbl(ws.cells(PG_LBL, C_THEME), "KEYWORD")
    Call Lbl(ws.cells(PG_LBL, C_FILTER), "FILTER")
    ws.Range(ws.cells(PG_IN, C_TGT), ws.cells(PG_IN, C_STATUS)).Merge
    ws.Range(ws.cells(PG_IN, C_FILTER), ws.cells(PG_IN, C_FILTER + 1)).Merge
    Call InputBox_(ws.Range(ws.cells(PG_IN, C_TGT), ws.cells(PG_IN, C_STATUS)), q)
    Call InputBox_(ws.cells(PG_IN, C_THEME), kw)             ' one cell, not merged (user, 2026-09-16)
    Call InputBox_(ws.Range(ws.cells(PG_IN, C_FILTER), ws.cells(PG_IN, C_FILTER + 1)), fl)
    ws.cells(PG_LBL, C_TGT).AddComment "Tickers, comma = several.  @PORT = RR4 position log, @WATCH = WATCHLIST"
    ws.cells(PG_LBL, C_FILTER).AddComment "Space / comma separated, any one matching keeps the note (OR):" & vbLf & _
        "STATUS words, GREEN, RED, MOAT RISK CATALYST, tiers T0 T1 T1L T2 T3 D," & vbLf & ">2026-08-01, <2026-06-30, 2026-08 or 2026 (call-date prefix)"
    ws.Rows(PG_IN).RowHeight = 20

    ' ---- data ----
    Dim lo As ListObject: Set lo = NotesTable()
    Dim n As Long
    Dim tg() As String, ty() As String, dt() As Date, st() As String, ro() As String, th() As String, be() As String, ev() As String
    Dim rw() As Long                                        ' rw(i) = note i's tblNotes body row (for the reading panel)
    ' ev(i, t) = note i's evidence in tier t (1..NTIER); at most one tier is filled
    n = ReadNotes(lo, tg, ty, dt, st, ro, th, be, ev, rw)

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
    For i = 1 To n                                          ' a macro target with any INDUSTRY note sits under INDUSTRY
        If ty(i) = "industry" Then
            k = UCase$(tg(i))
            If kind(k) = "macro" Then kind(k) = "industry"
        End If
    Next i
    With ws.cells(PG_TOTAL, 1)
        .Value = Format(n, "#,##0") & " notes . " & calls.count & " calls . " & tickers.count & " tickers" & _
                 IIf(indexMode, "      (index: double-click a target = query it; date grey > " & FRESH_GREY & "d, orange > " & FRESH_ORANGE & "d)", _
                                "      (double-click a note = full text, a ticker = its archived thesis, the title = close)")
        .Font.Color = CLR_SOFT: .Font.Bold = True
    End With
    With ws.Range(ws.cells(PG_TOTAL, 1), ws.cells(PG_TOTAL, C_LAST)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE
    End With

    ' ---- index: nothing typed -> every target with its note count and latest call ----
    If indexMode Then
        Dim perTarget As Object: Set perTarget = CreateObject("Scripting.Dictionary")
        For i = 1 To n: perTarget(UCase$(tg(i))) = perTarget(UCase$(tg(i))) + 1: Next i
        Call DrawIndex(ws, perTarget, latest, kind, mode, n)
        Call EnsurePanel(ws)
        Call NavAdd(ws, "L")
        Application.ScreenUpdating = prevScr
        Application.EnableEvents = prevEv
        Exit Sub
    End If

    ' ---- filter: latest call per target; with a QUERY every call of the queried
    '      targets is drawn (newest first); KEYWORD and FILTER narrow either way ----
    Dim qs As Variant, kws As Variant, fs As Variant
    qs = ExpandQuery(SplitList(q)): kws = SplitList(kw): fs = SplitTokens(fl)
    Dim keep() As Boolean: ReDim keep(0 To n)
    Dim shown As Object: Set shown = CreateObject("Scripting.Dictionary")   ' target -> note count
    Dim shownCalls As Object: Set shownCalls = CreateObject("Scripting.Dictionary")
    Dim drawnNotes As Long
    For i = 1 To n
        k = UCase$(tg(i))
        If IsEmpty(qs) Then
            keep(i) = (dt(i) = latest(k))
        Else
            keep(i) = MatchesQuery(k, qs)
        End If
        If keep(i) And Not IsEmpty(kws) Then keep(i) = MatchesKeyword(th(i) & " " & be(i) & " " & AllTiers(ev, i), kws)
        If keep(i) And Not IsEmpty(fs) Then keep(i) = MatchesFilter(st(i), ro(i), dt(i), ev, i, fs)
        If keep(i) Then
            shown(k) = shown(k) + 1
            shownCalls(k & "|" & CLng(dt(i))) = True
            drawnNotes = drawnNotes + 1
        End If
    Next i
    With ws.cells(PG_IN, C_COUNT)
        .Value = shownCalls.count & " calls drawn . " & drawnNotes & " notes"
        .Font.Color = RGB(80, 200, 120): .Font.Bold = True
    End With

    ' ---- header ----
    ws.cells(PG_HDR, C_TGT).Value = IIf(mode = "ticker", "TICKER A-Z", "LATEST CALL")
    Dim ns As String: ns = NoteSort(ws)
    Dim mark As String
    If Right$(ns, 3) = "asc" Then mark = " " & ChrW(&H25BC) Else mark = " " & ChrW(&H25B2)
    ws.cells(PG_HDR, C_STATUS).Value = "STATUS" & IIf(Left$(ns, 6) = "status", mark, "")
    ws.cells(PG_HDR, C_ROLE).Value = "ROLE" & IIf(Left$(ns, 4) = "role", mark, "")
    ws.cells(PG_HDR, C_STATUS).AddComment "Double-click: sort notes in every block by STATUS (green > red), again = reversed, again = off"
    ws.cells(PG_HDR, C_ROLE).AddComment "Double-click: sort notes in every block by ROLE (MOAT > RISK > CATALYST), again = reversed, again = off"
    ws.cells(PG_HDR, C_TIER).Value = "TIER"
    ws.cells(PG_HDR, C_THEME).Value = "THEME"
    ws.cells(PG_HDR, C_BEHAV).Value = "BEHAVIOR"
    With ws.Range(ws.cells(PG_HDR, 1), ws.cells(PG_HDR, C_LAST))
        .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With
    ws.cells(PG_HDR, C_TIER).AddComment "Which evidence tier this note's evidence sits in:" & vbLf & _
        "T0 Q&A > T1 CALL > T1L FILING (lag info) > T2 RESEARCH > T3 MEDIA > " & ChrW(&H394) & " OWN (own derivation)." & vbLf & _
        "Double-click the note row to read the full evidence text in the panel."
    ws.cells(PG_HDR, C_TGT).AddComment "Double-click: sort blocks by latest call date <-> ticker A-Z"

    ' ---- blocks ----
    Dim r As Long: r = PG_LIST
    Dim sec As Variant
    For Each sec In Array("stock", "industry", "macro")
        Dim order As Variant: order = OrderedTargets(shown, latest, kind, CStr(sec), mode)
        If Not IsEmpty(order) Then
            Call SectionBanner(ws.Range(ws.cells(r, 1), ws.cells(r, C_LAST)))
            ws.cells(r, 1).Value = UCase$(CStr(sec))
            r = r + 2
            Dim ti As Long
            For ti = LBound(order) To UBound(order)
                r = DrawBlock(ws, r, CStr(order(ti)), latest(order(ti)), n, tg, dt, st, ro, th, be, ev, rw, keep, ns) + 1
            Next ti
        End If
    Next sec
    If r > PG_LIST Then                                     ' 2026-09-16: fixed three-line rows, text beyond that is clipped
        ws.Range(ws.cells(PG_LIST, C_THEME), ws.cells(r, C_LAST)).WrapText = True
        ws.Range(ws.cells(PG_LIST, 1), ws.cells(r, C_LAST)).VerticalAlignment = xlTop
        Dim noteRow As Long
        For noteRow = PG_LIST To r
            If CStr(ws.cells(noteRow, C_KEY).Value) <> "" Then ws.Rows(noteRow).RowHeight = NOTE_ROW_H
        Next noteRow
    End If
    If shown.count = 0 Then
        If n = 0 Then
            ws.cells(r, 1).Value = "No notes yet - add them on the ThesisNotes sheet."
        Else
            ws.cells(r, 1).Value = "Nothing matches QUERY / KEYWORD / FILTER." & NearestTargets(q, tickers)
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
        tg() As String, dt() As Date, st() As String, ro() As String, th() As String, be() As String, ev() As String, rw() As Long, _
        keep() As Boolean, ByVal noteSort As String) As Long
    Dim top As Long: top = r
    Dim i As Long, shownName As String
    ' every kept note of this target, newest call first (stable: sheet order within one date)
    Dim idx() As Long, m As Long, a As Long, b As Long, t As Long
    ReDim idx(1 To n + 1)
    For i = 1 To n
        If keep(i) And UCase$(tg(i)) = key Then
            m = m + 1: idx(m) = i
            b = m
            If noteSort <> "" Then b = 1                    ' note sort on: ties keep plain data-sheet order, not date order
            Do While b > 1
                If dt(idx(b - 1)) >= dt(idx(b)) Then Exit Do
                t = idx(b - 1): idx(b - 1) = idx(b): idx(b) = t
                b = b - 1
            Loop
        End If
    Next i
    If noteSort <> "" Then                                 ' note sort on: stable insertion sort by rank over the whole block
        Dim rk() As Long: ReDim rk(1 To m + 1)
        For a = 1 To m: rk(a) = NoteRank(noteSort, st(idx(a)), ro(idx(a))): Next a
        For a = 2 To m
            b = a
            Do While b > 1
                If rk(b - 1) <= rk(b) Then Exit Do
                t = idx(b - 1): idx(b - 1) = idx(b): idx(b) = t
                t = rk(b - 1): rk(b - 1) = rk(b): rk(b) = t
                b = b - 1
            Loop
        Next a
    End If
    Dim prevDt As Date, firstGroup As Boolean: firstGroup = True
    For a = 1 To m
        i = idx(a)
        If shownName = "" Then shownName = tg(i)
        If a > 1 And dt(i) <> prevDt And noteSort = "" Then    ' a new (older) call starts: label it on its first row
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
            sc.Value = UCase$(Dash(st(i))): sc.Font.Color = StatusColor(st(i))
            With ws.cells(r, C_ROLE)
                .Value = Dash(ro(i)): .Font.Color = RoleColor(ro(i)): .HorizontalAlignment = xlCenter
            End With
            Call TextCell(ws.cells(r, C_THEME), th(i), RGB(245, 245, 245))
            Call TextCell(ws.cells(r, C_BEHAV), be(i), CLR_SOFT)
            Dim tt As Long, foundTier As Long: foundTier = 0    ' which tier (if any) has this note's evidence
            For tt = 1 To NTIER
                If Trim$(ev(i, tt)) <> "" Then foundTier = tt: Exit For
            Next tt
            With ws.cells(r, C_TIER)
                .HorizontalAlignment = xlCenter
                If foundTier > 0 Then
                    .Value = TierShort(foundTier): .Font.Color = TierColor(foundTier)
                Else
                    .Value = "-": .Font.Color = CLR_MUTED
                End If
            End With
            ws.cells(r, C_KEY).Value = key
            ws.cells(r, C_ROW).Value = rw(i)
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
        ro() As String, th() As String, be() As String, ev() As String, rw() As Long) As Long
    ReDim tg(0 To 0): ReDim ty(0 To 0): ReDim dt(0 To 0): ReDim st(0 To 0)
    ReDim ro(0 To 0): ReDim th(0 To 0): ReDim be(0 To 0): ReDim ev(0 To 0, 1 To NTIER): ReDim rw(0 To 0)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = lo.DataBodyRange.Value
    Dim rows As Long: rows = UBound(v, 1)
    ReDim tg(0 To rows): ReDim ty(0 To rows): ReDim dt(0 To rows): ReDim st(0 To rows)
    ReDim ro(0 To rows): ReDim th(0 To rows): ReDim be(0 To rows): ReDim ev(0 To rows, 1 To NTIER): ReDim rw(0 To rows)
    Dim ncol As Long: ncol = UBound(v, 2)
    Dim i As Long, n As Long
    For i = 1 To rows
        If Trim$(CStr(v(i, NT_TARGET))) <> "" Then
            n = n + 1
            rw(n) = i
            tg(n) = Trim$(CStr(v(i, NT_TARGET)))
            ty(n) = LCase$(Trim$(CStr(v(i, NT_TYPE))))
            If ty(n) <> "macro" Then ty(n) = "stock"
            If ty(n) = "macro" And ncol >= NT_SUBTYPE Then          ' 2026-09-19: macro splits into industry / macro
                If LCase$(Trim$(CStr(v(i, NT_SUBTYPE)))) = "industry" Then ty(n) = "industry"
            End If
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
'  2026-09-16: index, FILTER, @PORT / @WATCH, note reading panel
' ================================================================
' Section banner (STOCK / INDUSTRY / MACRO): black like the page, orange bold
' text, dark grey rule underneath (2026-09-19, was a grey CLR_BANNER fill).
Private Sub SectionBanner(rng As Range)
    With rng
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RR4_ACCENT: .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RR4_LINE
    End With
End Sub

' Empty state: every target as (target . notes . latest call) in IDX_GROUPS
' column groups, STOCK / INDUSTRY / MACRO, always A-Z, no header row.
Private Sub DrawIndex(ws As Worksheet, perTarget As Object, latest As Object, kind As Object, ByVal mode As String, ByVal n As Long)
    Dim r As Long: r = PG_LIST
    Dim g As Long
    If n = 0 Then
        ws.cells(r, 1).Value = "No notes yet - add them on the ThesisNotes sheet."
        ws.cells(r, 1).Font.Color = CLR_MUTED
        Exit Sub
    End If
    ' 2026-09-19 (user): no header row and no sort toggle - the index row PG_HDR
    ' stays blank and targets are always A-Z; banners show the section name only.
    Dim sec As Variant
    For Each sec In Array("stock", "industry", "macro")
        Dim order As Variant: order = OrderedTargets(perTarget, latest, kind, CStr(sec), "ticker")
        If Not IsEmpty(order) Then
            Call SectionBanner(ws.Range(ws.cells(r, 1), ws.cells(r, IDX_GROUPS * 3)))
            ws.cells(r, 1).Value = UCase$(CStr(sec))
            r = r + 2
            Dim cnt As Long: cnt = UBound(order) - LBound(order) + 1
            Dim perCol As Long: perCol = (cnt + IDX_GROUPS - 1) \ IDX_GROUPS   ' fill down, then across
            Dim ti As Long
            For ti = 0 To cnt - 1
                Dim rr As Long: rr = r + (ti Mod perCol)
                Dim cc As Long: cc = 1 + (ti \ perCol) * 3
                Dim key As String: key = CStr(order(LBound(order) + ti))
                Dim age As Long: age = Date - CDate(latest(key))
                Dim clr As Long
                If age > FRESH_ORANGE Then
                    clr = RGB(230, 140, 40)
                ElseIf age > FRESH_GREY Then
                    clr = CLR_MUTED
                Else
                    clr = CLR_SOFT
                End If
                With ws.cells(rr, cc)
                    .NumberFormat = "@"
                    .Value = key
                    On Error Resume Next
                    .Errors(xlNumberAsText).Ignore = True
                    On Error GoTo 0
                    .Font.Name = ZH_FONT: .Font.Bold = True: .Font.Size = 10
                    .Font.Color = IIf(age > FRESH_ORANGE, clr, RGB(245, 245, 245))
                    .HorizontalAlignment = xlLeft
                End With
                With ws.cells(rr, cc + 1)
                    .Value = perTarget(key): .Font.Color = CLR_SOFT: .HorizontalAlignment = xlRight
                End With
                Call DateCell(ws.cells(rr, cc + 2), CDate(latest(key)))
                ws.cells(rr, cc + 2).Font.Color = clr
                ws.cells(rr, C_KEY).Value = key                 ' marks the row as index data for the double-click
            Next ti
            r = r + perCol + 1
        End If
    Next sec
End Sub

' Space / comma separated tokens, upper-cased.
Private Function SplitTokens(ByVal s As String) As Variant
    SplitTokens = SplitList(Replace(Replace(s, " ", ","), vbTab, ","))
End Function

' FILTER: the note passes when ANY token matches (user: union).
Private Function MatchesFilter(ByVal status As String, ByVal role As String, ByVal d As Date, ev() As String, ByVal i As Long, fs As Variant) As Boolean
    Dim w As Variant, tok As String, t As Long
    Dim tier As String
    For t = 1 To NTIER
        If Trim$(ev(i, t)) <> "" Then tier = UCase$(Split(TierName(t, False), " ")(0)): Exit For
    Next t
    Dim su As String: su = UCase$(Trim$(status))
    Dim ru As String: ru = UCase$(Trim$(role))
    Dim isGreen As Boolean: isGreen = (su = "ROBUST" Or su = "SOLID" Or su = "GROWING")
    Dim isRed As Boolean: isRed = (su = "SLOWING" Or su = "SLUGGISH" Or su = "CHALLENGING" Or su = "CONTRACTION" Or su = "WARNING")
    Dim ds As String: ds = Format$(d, "yyyy-mm-dd")
    For Each w In fs
        tok = CStr(w)
        Select Case True
            Case tok = "GREEN": If isGreen Then MatchesFilter = True
            Case tok = "RED": If isRed Then MatchesFilter = True
            Case tok = su And su <> "": MatchesFilter = True
            Case tok = ru And ru <> "": MatchesFilter = True
            Case tok = "DELTA" Or tok = ChrW(&H394): If tier = "D" Then MatchesFilter = True
            Case tok = tier And tier <> "": MatchesFilter = True
            Case Left$(tok, 2) = ">=": If ds >= NormDate(Mid$(tok, 3)) Then MatchesFilter = True
            Case Left$(tok, 2) = "<=": If ds <= NormDate(Mid$(tok, 3)) Then MatchesFilter = True
            Case Left$(tok, 1) = ">": If ds > NormDate(Mid$(tok, 2)) Then MatchesFilter = True
            Case Left$(tok, 1) = "<": If ds < NormDate(Mid$(tok, 2)) Then MatchesFilter = True
            Case Len(tok) >= 4 And IsNumeric(Left$(tok, 4)): If Left$(ds, Len(tok)) = NormDate(tok) Then MatchesFilter = True
        End Select
        If MatchesFilter Then Exit Function
    Next w
End Function

' "2026/8/1" / "2026-08-01" / "2026-08" -> "2026-08-01" / "2026-08" (zero-padded, dashes)
Private Function NormDate(ByVal s As String) As String
    Dim p As Variant: p = Split(Replace(s, "/", "-"), "-")
    Dim i As Long, out As String
    For i = 0 To UBound(p)
        If i = 0 Then out = p(i) Else out = out & "-" & Right$("0" & p(i), 2)
    Next i
    NormDate = out
End Function

' QUERY tokens with @PORT / @WATCH expanded to the RR4 page's tickers (bare, no .TW).
Private Function ExpandQuery(qs As Variant) As Variant
    ExpandQuery = qs
    If IsEmpty(qs) Then Exit Function
    Dim out As Object: Set out = CreateObject("Scripting.Dictionary")
    Dim q As Variant, t As Variant
    For Each q In qs
        If q = "@PORT" Or q = "@WATCH" Then
            For Each t In RR4Tickers(CStr(q) = "@PORT")
                out(BareTicker(CStr(t))) = True
            Next t
        ElseIf Trim$(CStr(q)) <> "" Then
            out(BareTicker(CStr(q))) = True
        End If
    Next q
    If out.count = 0 Then ExpandQuery = Empty Else ExpandQuery = out.keys
End Function

Private Function BareTicker(ByVal s As String) As String
    BareTicker = UCase$(Trim$(s))
    If Right$(BareTicker, 4) = ".TWO" Then BareTicker = Left$(BareTicker, Len(BareTicker) - 4)
    If Right$(BareTicker, 3) = ".TW" Then BareTicker = Left$(BareTicker, Len(BareTicker) - 3)
End Function

' Tickers typed on the RR4 page: the position log (B44 down, until blank) or
' the WATCHLIST saved rows (B29:B39).  Read straight off the sheet so this
' module does not depend on PortfolioDashboard's private readers.
Private Function RR4Tickers(ByVal positions As Boolean) As Variant
    Dim out As Object: Set out = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("RR4")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Dim r As Long, tk As String
        If positions Then
            r = RR4_POS_FIRST
            Do While r < RR4_POS_FIRST + 200
                tk = Trim$(CStr(ws.cells(r, RR4_LEFT + 1).Value))
                If tk = "" Then Exit Do
                out(UCase$(tk)) = True
                r = r + 1
            Loop
        Else
            For r = RR4_WL_FIRST To RR4_WL_LAST
                tk = Trim$(CStr(ws.cells(r, RR4_LEFT + 1).Value))
                If tk <> "" Then out(UCase$(tk)) = True
            Next r
        End If
    End If
    RR4Tickers = out.keys
End Function

' "Nothing matches" helper: targets starting with any typed QUERY token.
Private Function NearestTargets(ByVal q As String, tickers As Object) As String
    If Trim$(q) = "" Then Exit Function
    Dim qs As Variant: qs = SplitList(q)
    If IsEmpty(qs) Then Exit Function
    Dim k As Variant, w As Variant, hits As String, nHits As Long
    For Each k In tickers.keys
        For Each w In qs
            If Left$(CStr(w), 1) <> "@" And Len(CStr(w)) >= 2 Then
                If Left$(CStr(k), Len(CStr(w))) = CStr(w) Or InStr(CStr(k), CStr(w)) > 0 Then
                    hits = hits & IIf(hits = "", "", ", ") & CStr(k): nHits = nHits + 1
                    Exit For
                End If
            End If
        Next w
        If nHits >= 8 Then Exit For
    Next k
    If hits <> "" Then NearestTargets = "   Did you mean: " & hits
End Function

' Anchor TH_PANEL level with the row the user actually double-clicked (2026-09-20)
' instead of always the fixed header row: on a long, scrolled list the fixed
' anchor could sit above whatever part of the sheet was currently in view,
' making the panel look like it "disappeared" upward. Clamped at the header row
' so it can never climb above the inputs / nav bar either.
Private Sub PositionPanelAt(ws As Worksheet, shp As Shape, ByVal clickRow As Long)
    Dim minTop As Double: minTop = ws.Rows(PG_HDR + NavOffset(ws)).Top
    Dim want As Double
    If clickRow >= PG_HDR + NavOffset(ws) Then
        want = ws.Rows(clickRow).Top
    Else
        want = minTop
    End If
    If want < minTop Then want = minTop
    shp.Top = want
End Sub

' Reading panel for one note (tblNotes body row), plus the target's archived
' one-liner (S1) and next-check section (S6) as a reminder of the thesis.
Private Sub ShowNote(ws As Worksheet, ByVal key As String, ByVal bodyRow As Long, ByVal clickRow As Long)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("TH_PANEL")
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = NotesTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If bodyRow < 1 Or bodyRow > lo.ListRows.count Then Exit Sub
    Dim v As Variant: v = lo.DataBodyRange.Rows(bodyRow).Value
    shp.Left = ws.Columns(C_BEHAV + NavLeft(ws)).Left
    Call PositionPanelAt(ws, shp, clickRow)
    shp.Visible = msoTrue
    Dim tr As Object: Set tr = shp.TextFrame2.TextRange

    Dim tier As Long, t As Long
    For t = 1 To NTIER
        If Trim$(CStr(v(1, NT_T0 + t - 1))) <> "" Then tier = t: Exit For
    Next t
    Dim head As String: head = CStr(v(1, NT_TARGET)) & "   " & CStr(v(1, NT_THEME))
    Dim meta As String
    meta = Format$(v(1, NT_DATE), "yyyy/mm/dd") & "  .  " & UCase$(Dash(CStr(v(1, NT_STATUS)))) & "  .  " & Dash(CStr(v(1, NT_ROLE))) & _
           "  .  " & IIf(tier > 0, TierName(tier, True), "no evidence")
    Dim txt As String
    Dim starts(0 To 5) As Long, lens(0 To 5) As Long
    starts(0) = 1: lens(0) = Len(head): txt = head & vbCr
    starts(1) = Len(txt) + 1: lens(1) = Len(meta): txt = txt & meta & vbCr & vbCr
    Dim h As String
    h = "BEHAVIOR": starts(2) = Len(txt) + 1: lens(2) = Len(h)
    txt = txt & h & vbCr & Dash(CStr(v(1, NT_BEHAV))) & vbCr & vbCr
    h = IIf(tier > 0, TierName(tier, True), "EVIDENCE"): starts(3) = Len(txt) + 1: lens(3) = Len(h)
    txt = txt & h & vbCr & IIf(tier > 0, CStr(v(1, NT_T0 + tier - 1)), "-") & vbCr & vbCr

    ' archived thesis reminder
    Dim la As ListObject, row As Long
    On Error Resume Next
    Set la = ThisWorkbook.Worksheets(ARCHIVE_SHEET).ListObjects(THESIS_TABLE)
    On Error GoTo 0
    If Not la Is Nothing Then row = ArchiveRow(la, key)
    h = "ARCHIVE  " & L("S1"): starts(4) = Len(txt) + 1: lens(4) = Len(h)
    If row = 0 Then
        txt = txt & h & vbCr & "(no archived thesis for " & key & ")" & vbCr
        starts(5) = Len(txt): lens(5) = 0
    Else
        Dim b As Range: Set b = la.DataBodyRange
        txt = txt & h & vbCr & Dash(Trim$(CStr(b.cells(row, 6).Value))) & vbCr & vbCr
        h = "ARCHIVE  " & L("S6"): starts(5) = Len(txt) + 1: lens(5) = Len(h)
        txt = txt & h & vbCr & Dash(Trim$(CStr(b.cells(row, 11).Value))) & vbCr
    End If
    txt = txt & vbCr & "(click this panel to close)"

    tr.Text = txt
    tr.Font.Name = ZH_FONT: tr.Font.NameFarEast = ZH_FONT
    tr.Font.Size = 10: tr.Font.Bold = msoFalse
    tr.Font.Fill.ForeColor.RGB = RGB(215, 215, 215)
    tr.ParagraphFormat.SpaceAfter = 2
    With tr.Characters(starts(0), lens(0)).Font
        .Size = 12: .Bold = msoTrue: .Fill.ForeColor.RGB = RR4_ACCENT
    End With
    With tr.Characters(starts(1), lens(1)).Font
        .Size = 9: .Fill.ForeColor.RGB = RGB(140, 140, 140)
    End With
    Dim k As Long
    For k = 2 To 5
        If lens(k) > 0 Then
            With tr.Characters(starts(k), lens(k)).Font
                .Bold = msoTrue: .Size = 10
                .Fill.ForeColor.RGB = IIf(k >= 4, RGB(140, 140, 140), RR4_ACCENT)
            End With
        End If
    Next k
    Dim tail As Long: tail = Len(txt) - Len("(click this panel to close)")
    With tr.Characters(tail + 1, Len(txt) - tail).Font
        .Size = 8: .Fill.ForeColor.RGB = RGB(110, 110, 110)
    End With
End Sub

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
    shp.OnAction = "modThesis.ThesisPanelClick"      ' a click on the panel closes it (and cannot enter text-edit)
    shp.Visible = msoFalse                           ' shown on demand by ShowArchive / ShowNote
End Sub

' Click on TH_PANEL (its OnAction): hide it.  Double-clicking the page title still works too.
Public Sub ThesisPanelClick()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(THESIS_SHEET)
    ws.Shapes("TH_PANEL").Visible = msoFalse
End Sub

Private Sub ShowArchive(ws As Worksheet, ByVal key As String, Optional ByVal clickRow As Long = 0)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("TH_PANEL")
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    If key = "" Then shp.Visible = msoFalse: Exit Sub
    ' float over BEHAVIOR, level with the clicked row (a reading overlay)
    shp.Left = ws.Columns(C_BEHAV + NavLeft(ws)).Left
    Call PositionPanelAt(ws, shp, clickRow)
    shp.Visible = msoTrue
    Dim tr As Object: Set tr = shp.TextFrame2.TextRange

    Dim lo As ListObject, row As Long
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(ARCHIVE_SHEET).ListObjects(THESIS_TABLE)
    On Error GoTo 0
    If Not lo Is Nothing Then row = ArchiveRow(lo, key)

    If row = 0 Then
        tr.Text = key & vbCr & "No archived thesis for this target - its notes are all there is." & vbCr & _
                  "(click this panel to close)"
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
