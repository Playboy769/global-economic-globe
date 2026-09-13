Attribute VB_Name = "modThesis"
Option Explicit

' ================================================================
'  THESIS LIBRARY (nav code TH, TH! = (re)format) - sheet "Thesis Library"
' ----------------------------------------------------------------
'  2026-09-13.  A place to keep macro and single-stock investment theses
'  and read them quickly.  One row per thesis in the ListObject tblThesis:
'
'    A date | B type (macro/stock) | C target | D title | E stance |
'    F..K the six sections of the 500-word thesis checklist
'         (research/_templates/500-word-investment-thesis-checklist-template.md):
'         1 one-line thesis  2 business model & moat  3 why now
'         4 three supporting numbers  5 biggest risk & falsifiers
'         6 decision & next checkpoint |
'    L next check date | M status (ACTIVE / CONFIRMED / FALSIFIED / CLOSED)
'
'  New theses are typed straight into the row under the table (the
'  ListObject grows; type / stance / status carry dropdowns).  Double-
'  clicking a row shows its six sections in the reading panel (text box TH_PANEL to
'  the right of the table) - see RR4/SheetThesis_Code.txt, written into the
'  sheet module by EnsureThesisSheetCode.  Chinese text uses Noto Sans TC;
'  the labels are ChrW-built so this file stays ASCII (VBE import rule).
'
'  TH! never touches the data: it (re)builds the header, widths, dropdowns,
'  the panel and the nav bar around whatever rows are already there.
'
'  Sorting: double-click a header of a short column (date / type / target /
'  title / stance / next check / status) to sort by it; again = reverse.
'  First click is descending (newest first); stance and status sort in their
'  own order (LONG WATCH NEUTRAL SHORT / ACTIVE CONFIRMED FALSIFIED CLOSED)
'  first.  The header shows an arrow; the sort is kept in the hidden sheet name
'  THSORT ("col|dir") and re-applied by TH! and when a target is typed.
'  Double-clicking the page title drops the sort and goes back to date
'  oldest -> newest (no input-order column is kept).
' ================================================================

Public Const THESIS_SHEET As String = "Thesis Library"
Public Const THESIS_TABLE As String = "tblThesis"
Private Const TH_HDR As Long = 4                 ' page row of the header (title row 2, row 3 blank); bar block adds 4
Private Const TH_LEFT As Long = 0                ' draw from column A; NavAdd inserts the blank column A afterwards (modNav)
Private Const TH_NCOL As Long = 13
Private Const PANEL_COL As Long = 15             ' O at draw time (table A:M, N the gap); P once NavAdd adds column A
Private Const PANEL_W As Double = 620
Private Const PANEL_H As Double = 760
Private Const ZH_FONT As String = "Noto Sans TC"
Private Const DEL_WINDOW As Double = 2#          ' seconds: a second double-click on the same row within this deletes it
Private Const SORT_MARK As String = "THSORT"     ' "col|dir" of the current sort (dir 2 = first click, 1 = reversed)
Private m_lastDblRow As Long                     ' row of the last double-click (0 = none)
Private m_lastDblAt As Double                    ' Timer of that double-click

Private Function L(ByVal key As String) As String
    Select Case key
        Case "DATE": L = ChrW(&H65E5) & ChrW(&H671F)
        Case "TYPE": L = ChrW(&H985E) & ChrW(&H578B)
        Case "TARGET": L = ChrW(&H6A19) & ChrW(&H7684)
        Case "TITLE": L = ChrW(&H6A19) & ChrW(&H984C)
        Case "STANCE": L = ChrW(&H7ACB) & ChrW(&H5834)
        Case "S1": L = ChrW(&H4E00) & ChrW(&H53E5) & ChrW(&H8A71) & ChrW(&H8AD6) & ChrW(&H9EDE)
        Case "S2": L = ChrW(&H5546) & ChrW(&H696D) & ChrW(&H6A21) & ChrW(&H5F0F) & ChrW(&H8207) & ChrW(&H8B77) & ChrW(&H57CE) & ChrW(&H6CB3)
        Case "S3": L = ChrW(&H70BA) & ChrW(&H4EC0) & ChrW(&H9EBC) & ChrW(&H662F) & ChrW(&H73FE) & ChrW(&H5728)
        Case "S4": L = ChrW(&H4E09) & ChrW(&H500B) & ChrW(&H652F) & ChrW(&H6490) & ChrW(&H6578) & ChrW(&H5B57)
        Case "S5": L = ChrW(&H6700) & ChrW(&H5927) & ChrW(&H98A8) & ChrW(&H96AA) & ChrW(&H8207) & ChrW(&H8B49) & ChrW(&H507D) & ChrW(&H8A0A) & ChrW(&H865F)
        Case "S6": L = ChrW(&H6C7A) & ChrW(&H7B56) & ChrW(&H8207) & ChrW(&H4E0B) & ChrW(&H4E00) & ChrW(&H500B) & ChrW(&H9A57) & ChrW(&H8B49) & ChrW(&H9EDE)
        Case "NEXT": L = ChrW(&H4E0B) & ChrW(&H6B21) & ChrW(&H9A57) & ChrW(&H8B49) & ChrW(&H65E5)
        Case "STATUS": L = ChrW(&H72C0) & ChrW(&H614B)
        Case "PANEL_EMPTY": L = ChrW(&H96D9) & ChrW(&H64CA) & ChrW(&H5DE6) & ChrW(&H5074) & ChrW(&H4EFB) & ChrW(&H4E00) & ChrW(&H5217) & ChrW(&HFF0C) & ChrW(&H9019) & ChrW(&H88E1) & ChrW(&H986F) & ChrW(&H793A) & ChrW(&H8A72) & ChrW(&H7BC7) & ChrW(&H0020) & ChrW(&H0074) & ChrW(&H0068) & ChrW(&H0065) & ChrW(&H0073) & ChrW(&H0069) & ChrW(&H0073) & ChrW(&H0020) & ChrW(&H7684) & ChrW(&H516D) & ChrW(&H6BB5) & ChrW(&H5168) & ChrW(&H6587) & ChrW(&HFF1B) & ChrW(&H0032) & ChrW(&H0020) & ChrW(&H79D2) & ChrW(&H5167) & ChrW(&H518D) & ChrW(&H96D9) & ChrW(&H64CA) & ChrW(&H540C) & ChrW(&H4E00) & ChrW(&H5217) & ChrW(&HFF1D) & ChrW(&H522A) & ChrW(&H9664) & ChrW(&H8A72) & ChrW(&H7BC7) & ChrW(&HFF08) & ChrW(&H6703) & ChrW(&H5148) & ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&HFF09)
        Case "DEL_ASK": L = ChrW(&H522A) & ChrW(&H9664) & ChrW(&H9019) & ChrW(&H7BC7) & ChrW(&H0020) & ChrW(&H0074) & ChrW(&H0068) & ChrW(&H0065) & ChrW(&H0073) & ChrW(&H0069) & ChrW(&H0073) & ChrW(&HFF1F)
        Case "DEL_DONE": L = ChrW(&H5DF2) & ChrW(&H522A) & ChrW(&H9664)
        Case "TITLE_ZH": L = ChrW(&H7E3D) & ChrW(&H7D93) & ChrW(&H8207) & ChrW(&H500B) & ChrW(&H80A1) & ChrW(&H0020) & ChrW(&H0074) & ChrW(&H0068) & ChrW(&H0065) & ChrW(&H0073) & ChrW(&H0069) & ChrW(&H0073) & ChrW(&H0020) & ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H5EAB)
    End Select
End Function

' Column headers, in table order.
Private Function Headers() As Variant
    Headers = Array(L("DATE"), L("TYPE"), L("TARGET"), L("TITLE"), L("STANCE"), _
                    "1 " & L("S1"), "2 " & L("S2"), "3 " & L("S3"), "4 " & L("S4"), "5 " & L("S5"), "6 " & L("S6"), _
                    L("NEXT"), L("STATUS"))
End Function

' ----------------------------------------------------------------
Sub BuildThesisLibrary()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(THESIS_SHEET)
    On Error GoTo 0
    Dim isNew As Boolean
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = THESIS_SHEET
        isNew = True
    End If
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Call NavStrip(ws)

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(THESIS_TABLE)
    On Error GoTo 0

    ' ---- page look (data untouched: only formats) ----
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .VerticalAlignment = xlTop
    End With
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False

    With ws.cells(2, 1 + TH_LEFT)
        .Value = "THESIS LIBRARY"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 14
    End With
    ws.Rows(2).RowHeight = 24
    With ws.cells(2, 4 + TH_LEFT)
        .Value = L("TITLE_ZH") & "  .  one row per thesis, six sections of the 500-word checklist  .  formatted " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RGB(120, 120, 120): .Font.Size = 9: .Font.Name = ZH_FONT
    End With

    ' ---- table ----
    Dim hdr As Variant: hdr = Headers()
    Dim j As Long
    If lo Is Nothing Then
        For j = 0 To TH_NCOL - 1: ws.cells(TH_HDR, j + 1 + TH_LEFT).Value = hdr(j): Next j
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.cells(TH_HDR, 1 + TH_LEFT), ws.cells(TH_HDR + 1, TH_NCOL + TH_LEFT)), , xlYes)
        lo.Name = THESIS_TABLE
    Else
        For j = 0 To TH_NCOL - 1: lo.HeaderRowRange.cells(1, j + 1).Value = hdr(j): Next j
    End If
    lo.TableStyle = ""
    lo.ShowTableStyleRowStripes = False
    With lo.HeaderRowRange
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 9: .Font.Name = ZH_FONT
        .Interior.Color = RGB(0, 0, 0)
        .WrapText = False
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RR4_LINE
    End With
    Dim widths As Variant
    widths = Array(10, 7, 9, 26, 8, 30, 18, 18, 18, 18, 18, 10, 10)
    For j = 0 To TH_NCOL - 1: ws.Columns(j + 1 + TH_LEFT).ColumnWidth = widths(j): Next j
    ws.Columns(TH_NCOL + 1 + TH_LEFT).ColumnWidth = 2
    If Not lo.DataBodyRange Is Nothing Then Call FormatThesisRows(lo)
    Call ApplyThesisSort(ws, lo)                     ' re-sort + header arrow (headers were just rewritten)

    ' ---- dropdowns on the whole columns of the table (they extend with it) ----
    Call SetList(lo.ListColumns(2).DataBodyRange, "macro,stock")
    Call SetList(lo.ListColumns(5).DataBodyRange, "LONG,SHORT,NEUTRAL,WATCH")
    Call SetList(lo.ListColumns(13).DataBodyRange, "ACTIVE,CONFIRMED,FALSIFIED,CLOSED")

    ' ---- reading panel ----
    Call EnsurePanel(ws)
    Call ShowThesis(ws, Nothing)

    Call NavAdd(ws, "TH")
    ws.cells(TH_HDR + 1 + NavOffset(ws), 1 + TH_LEFT).Select
    Call EnsureThesisSheetCode(ws)
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    Call NavNotify("Thesis Library " & IIf(isNew, "created", "formatted") & ": " & ThesisCount(lo) & " theses")
    Exit Sub
Fail:
    Dim msg As String: msg = Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    On Error Resume Next
    Call NavAdd(ws, "TH")
    Call NavNotify("Thesis Library failed: " & msg, True)
End Sub

Private Function ThesisCount(lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim r As Range, n As Long
    For Each r In lo.ListColumns(3).DataBodyRange.cells
        If Trim(CStr(r.Value)) <> "" Then n = n + 1
    Next r
    ThesisCount = n
End Function

' Row look: dates / short columns Consolas, the text columns Noto Sans TC,
' long text clipped (not wrapped - the panel is for reading).
Public Sub FormatThesisRows(lo As ListObject)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    With lo.DataBodyRange
        .Interior.Color = RGB(8, 8, 8)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas": .Font.Size = 9
        .WrapText = False
        .VerticalAlignment = xlCenter
        .RowHeight = 18
    End With
    lo.ListColumns(1).DataBodyRange.NumberFormat = "yyyy/mm/dd"
    lo.ListColumns(12).DataBodyRange.NumberFormat = "yyyy/mm/dd"
    Dim c As Long
    For c = 3 To 11
        lo.ListColumns(c).DataBodyRange.Font.Name = ZH_FONT
    Next c
    lo.ListColumns(3).DataBodyRange.Font.Color = RR4_ACCENT
    lo.ListColumns(3).DataBodyRange.Font.Bold = True
    Dim r As Range
    For Each r In lo.ListColumns(5).DataBodyRange.cells
        r.Font.Color = StanceColor(CStr(r.Value)): r.Font.Bold = True
    Next r
    For Each r In lo.ListColumns(13).DataBodyRange.cells
        r.Font.Color = StatusColor(CStr(r.Value)): r.Font.Bold = True
    Next r
End Sub

Private Function StanceColor(ByVal s As String) As Long
    Select Case UCase(s)
        Case "LONG":    StanceColor = RGB(220, 80, 80)       ' TW convention: red = up
        Case "SHORT":   StanceColor = RGB(80, 200, 120)
        Case "WATCH":   StanceColor = RGB(249, 168, 37)
        Case Else:      StanceColor = RGB(180, 180, 180)
    End Select
End Function

Private Function StatusColor(ByVal s As String) As Long
    Select Case UCase(s)
        Case "ACTIVE":    StatusColor = RGB(221, 221, 221)
        Case "CONFIRMED": StatusColor = RGB(80, 200, 120)
        Case "FALSIFIED": StatusColor = RGB(220, 80, 80)
        Case "CLOSED":    StatusColor = RGB(110, 110, 110)
        Case Else:        StatusColor = RGB(180, 180, 180)
    End Select
End Function

Private Sub SetList(rng As Range, ByVal items As String)
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=items
    rng.Validation.ShowError = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' The reading panel: one text box right of the table.
Private Sub EnsurePanel(ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("TH_PANEL")
    On Error GoTo 0
    Dim topRow As Long: topRow = TH_HDR + NavOffset(ws)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, ws.Columns(PANEL_COL).Left, ws.Rows(topRow).Top, PANEL_W, PANEL_H)
        shp.Name = "TH_PANEL"
    Else
        shp.Left = ws.Columns(PANEL_COL).Left: shp.Top = ws.Rows(topRow).Top
        shp.Width = PANEL_W: shp.Height = PANEL_H
    End If
    shp.Placement = xlMove
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(12, 12, 12)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RR4_LINE
    shp.Line.Weight = 0.75
    With shp.TextFrame2
        .MarginLeft = 12: .MarginRight = 12: .MarginTop = 10: .MarginBottom = 10
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeShapeToFitText      ' a long thesis grows the panel instead of being clipped
        .VerticalAnchor = msoAnchorTop
    End With
End Sub

' Fill the panel from a table row (Nothing / a row outside the table ->
' the hint).  Called by the sheet's SelectionChange.
Public Sub ShowThesis(ws As Worksheet, Target As Range)
    Dim lc As Long: lc = NavLeft(ws)                 ' blank column A once the bar is on
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("TH_PANEL")
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(THESIS_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    Dim r As Long: r = 0
    If Not Target Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            If Not Intersect(Target.cells(1, 1), lo.DataBodyRange) Is Nothing Then r = Target.cells(1, 1).Row
        End If
    End If
    Dim tr As Object: Set tr = shp.TextFrame2.TextRange     ' TextRange2 late-bound: no Office type-library dependency
    ' VBA's Or does not short-circuit: cells(0, 3) would blow up, so test r first
    Dim blank As Boolean: blank = (r = 0)
    If Not blank Then blank = (Trim(CStr(ws.cells(r, 3 + lc).Value)) = "")
    If blank Then
        tr.Text = L("PANEL_EMPTY")
        tr.Font.Name = ZH_FONT: tr.Font.NameFarEast = ZH_FONT: tr.Font.Size = 10
        tr.Font.Fill.ForeColor.RGB = RGB(110, 110, 110): tr.Font.Bold = msoFalse
        Exit Sub
    End If

    ' ---- compose: head line, meta line, six sections ----
    Dim secKeys As Variant: secKeys = Array("S1", "S2", "S3", "S4", "S5", "S6")
    Dim txt As String, starts(0 To 7) As Long, lens(0 To 7) As Long
    Dim head As String: head = CStr(ws.cells(r, 3 + lc).Value) & "   " & CStr(ws.cells(r, 4 + lc).Value)
    starts(0) = 1: lens(0) = Len(head)
    txt = head & vbCr
    Dim meta As String
    meta = CStr(ws.cells(r, 2 + lc).Value) & "  .  " & CStr(ws.cells(r, 5 + lc).Value) & "  .  " & CStr(ws.cells(r, 13 + lc).Value) & _
           "  .  " & Format(ws.cells(r, 1 + lc).Value, "yyyy/mm/dd") & _
           IIf(ws.cells(r, 12 + lc).Value <> "", "  .  " & L("NEXT") & " " & Format(ws.cells(r, 12 + lc).Value, "yyyy/mm/dd"), "")
    starts(1) = Len(txt) + 1: lens(1) = Len(meta)
    txt = txt & meta & vbCr & vbCr
    Dim k As Long
    For k = 0 To 5
        Dim h As String: h = (k + 1) & "  " & L(CStr(secKeys(k)))
        starts(k + 2) = Len(txt) + 1: lens(k + 2) = Len(h)
        Dim body As String: body = Trim(CStr(ws.cells(r, 6 + k + lc).Value))
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

' Sheet events (RR4/SheetThesis_Code.txt): double-click a row -> panel; a
' new row gets today's date and ACTIVE when its target is typed.
Public Sub ThesisDoubleClick(ws As Worksheet, ByVal Target As Range, ByRef Cancel As Boolean)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(THESIS_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    Dim lc As Long: lc = NavLeft(ws)
    ' page title -> drop the sort; header -> sort by that column
    If Target.cells(1, 1).Row = 2 + NavOffset(ws) And Target.cells(1, 1).Column = 1 + lc Then
        Cancel = True
        Call ClearThesisSort(ws, lo)
        Exit Sub
    End If
    If Not Intersect(Target.cells(1, 1), lo.HeaderRowRange) Is Nothing Then
        Cancel = True
        Call ThesisSortClick(ws, lo, Target.cells(1, 1).Column - lo.HeaderRowRange.Column + 1)
        Exit Sub
    End If
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target.cells(1, 1), lo.DataBodyRange) Is Nothing Then Exit Sub
    Cancel = True                                    ' no in-cell edit on a double-click
    Dim r As Long: r = Target.cells(1, 1).Row
    ' Excel has no triple-click event (click 3 is just a click), so "double-
    ' click again within DEL_WINDOW seconds on the same row" is the delete gesture
    Dim el As Double: el = Timer - m_lastDblAt: If el < 0 Then el = el + 86400
    If r = m_lastDblRow And el <= DEL_WINDOW Then
        m_lastDblRow = 0
        Dim nm As String: nm = CStr(ws.cells(r, 3 + lc).Value) & "  " & CStr(ws.cells(r, 4 + lc).Value)
        If MsgBox(L("DEL_ASK") & vbCrLf & vbCrLf & nm, vbYesNo + vbQuestion, "Thesis Library") = vbYes Then
            Dim prevEv As Boolean: prevEv = Application.EnableEvents
            Application.EnableEvents = False
            On Error Resume Next
            lo.ListRows(r - lo.HeaderRowRange.Row).Delete
            On Error GoTo 0
            Application.EnableEvents = prevEv
            Call ShowThesis(ws, Nothing)
            Call NavNotify("Thesis Library: " & L("DEL_DONE") & " - " & nm)
        End If
        Exit Sub
    End If
    m_lastDblRow = r: m_lastDblAt = Timer
    Call ShowThesis(ws, Target)
End Sub

Public Sub ThesisChange(ws As Worksheet, ByVal Target As Range)
    Dim lc As Long: lc = NavLeft(ws)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(THESIS_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    Dim c As Range, typedTarget As Boolean
    For Each c In Intersect(Target, lo.ListColumns(3).DataBodyRange).cells
        If Trim(CStr(c.Value)) <> "" Then
            typedTarget = True
            If ws.cells(c.Row, 1 + lc).Value = "" Then ws.cells(c.Row, 1 + lc).Value = Date
            If ws.cells(c.Row, 13 + lc).Value = "" Then ws.cells(c.Row, 13 + lc).Value = "ACTIVE"
        End If
    Next c
    If typedTarget Then Call ApplyThesisSort(ws, lo)   ' a new / renamed thesis falls into the kept sort
    Call FormatThesisRows(lo)
    Call SetList(lo.ListColumns(2).DataBodyRange, "macro,stock")
    Call SetList(lo.ListColumns(5).DataBodyRange, "LONG,SHORT,NEUTRAL,WATCH")
    Call SetList(lo.ListColumns(13).DataBodyRange, "ACTIVE,CONFIRMED,FALSIFIED,CLOSED")
    Application.EnableEvents = prevEv
End Sub

' ----------------------------------------------------------------
' Sorting (header double-click).  Long text columns 6-11 are not sortable.
Private Function IsSortableCol(ByVal col As Long) As Boolean
    Select Case col
        Case 1, 2, 3, 4, 5, 12, 13: IsSortableCol = True
    End Select
End Function

Private Sub ThesisSortClick(ws As Worksheet, lo As ListObject, ByVal col As Long)
    If Not IsSortableCol(col) Then Exit Sub
    Dim curCol As Long, curDir As Long
    Call ReadSortMark(ws, curCol, curDir)
    Dim newDir As Long: newDir = 2
    If curCol = col Then newDir = 3 - curDir
    On Error Resume Next
    ws.Names(SORT_MARK).Delete
    On Error GoTo 0
    ws.Names.Add Name:=SORT_MARK, RefersTo:="=""" & col & "|" & newDir & """", Visible:=False
    m_lastDblRow = 0                                 ' rows moved: a follow-up double-click must not delete
    Call ApplyThesisSort(ws, lo)
    Call NavNotify("Thesis Library sorted by column " & col & IIf(newDir = 2, " (first order)", " (reversed)"))
End Sub

Private Sub ClearThesisSort(ws As Worksheet, lo As ListObject)
    On Error Resume Next
    ws.Names(SORT_MARK).Delete
    On Error GoTo 0
    m_lastDblRow = 0
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    If Not lo.DataBodyRange Is Nothing Then
        Call SortTable(lo, 1, xlAscending, "")       ' date oldest -> newest
        Call FormatThesisRows(lo)
    End If
    Call WriteHeaderArrows(lo, 0, 0)
    Application.EnableEvents = prevEv
    Call NavNotify("Thesis Library: sort cleared (date oldest first)")
End Sub

' Re-apply the kept sort (if any) and redraw the header arrows.
Public Sub ApplyThesisSort(ws As Worksheet, lo As ListObject)
    Dim col As Long, dir As Long
    Call ReadSortMark(ws, col, dir)
    If Not IsSortableCol(col) Then col = 0
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fin
    If col > 0 And Not lo.DataBodyRange Is Nothing Then
        Dim lst As String
        Select Case col
            Case 5: lst = "LONG,WATCH,NEUTRAL,SHORT"
            Case 13: lst = "ACTIVE,CONFIRMED,FALSIFIED,CLOSED"
        End Select
        ' first click (dir 2): descending, except the enum columns which start in their own order
        Dim ord As Long
        If lst <> "" Then
            ord = IIf(dir = 2, xlAscending, xlDescending)
        Else
            ord = IIf(dir = 2, xlDescending, xlAscending)
        End If
        Call SortTable(lo, col, ord, lst)
        Call FormatThesisRows(lo)
    End If
    Call WriteHeaderArrows(lo, col, dir)
Fin:
    Application.EnableEvents = prevEv
End Sub

Private Sub SortTable(lo As ListObject, ByVal col As Long, ByVal ord As Long, ByVal customList As String)
    With lo.Sort
        .SortFields.Clear
        If customList <> "" Then
            ' CVar: a plain String variable here raises error 13 (type mismatch); literals and CVar work
            .SortFields.Add Key:=lo.ListColumns(col).DataBodyRange, SortOn:=xlSortOnValues, Order:=ord, CustomOrder:=CVar(customList)
        Else
            .SortFields.Add Key:=lo.ListColumns(col).DataBodyRange, SortOn:=xlSortOnValues, Order:=ord
        End If
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

' Headers back to their labels; the sorted one gets a down (first order) or up (reversed) arrow.
Private Sub WriteHeaderArrows(lo As ListObject, ByVal col As Long, ByVal dir As Long)
    Dim hdr As Variant: hdr = Headers()
    Dim j As Long, s As String
    For j = 0 To TH_NCOL - 1
        s = hdr(j)
        If j + 1 = col Then s = s & " " & IIf(dir = 2, ChrW(&H25BC), ChrW(&H25B2))
        If CStr(lo.HeaderRowRange.cells(1, j + 1).Value) <> s Then lo.HeaderRowRange.cells(1, j + 1).Value = s
    Next j
End Sub

' Walk ws.Names instead of ws.Names(SORT_MARK): after a reopen the sheet-level
' name can fail to resolve by its short name (same as modNav.NavGeom).
Private Sub ReadSortMark(ws As Worksheet, ByRef col As Long, ByRef dir As Long)
    col = 0: dir = 2
    Dim nm As Name
    For Each nm In ws.Names
        If Right(nm.Name, Len(SORT_MARK) + 1) = "!" & SORT_MARK Then
            Dim v As String: v = Replace(Replace(nm.RefersTo, "=", ""), """", "")
            If InStr(v, "|") > 0 Then
                col = CLng(Val(Split(v, "|")(0))): dir = CLng(Val(Split(v, "|")(1)))
                If dir <> 1 Then dir = 2
            End If
            Exit Sub
        End If
    Next nm
End Sub

' Write the sheet's event code into its document module (same approach as
' RRG.EnsureSheetCode; needs "Trust access to the VBA project object model").
Private Sub EnsureThesisSheetCode(ws As Worksheet)
    On Error GoTo Skip
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 Then
            If comp.Properties("Name").Value = ws.Name Then
                Dim cm As Object: Set cm = comp.CodeModule
                If cm.CountOfLines > 0 Then
                    If InStr(cm.Lines(1, cm.CountOfLines), "ThesisDoubleClick") > 0 Then Exit Sub
                    cm.DeleteLines 1, cm.CountOfLines
                End If
                cm.AddFromString "Option Explicit" & vbCrLf & vbCrLf & _
                    "' Thesis Library: double-click a row to read it in the panel (see modThesis.bas)" & vbCrLf & _
                    "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & vbCrLf & _
                    "    Call ThesisDoubleClick(Me, Target, Cancel)" & vbCrLf & _
                    "End Sub" & vbCrLf & vbCrLf & _
                    "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
                    "    Call ThesisChange(Me, Target)" & vbCrLf & _
                    "End Sub" & vbCrLf
                Exit Sub
            End If
        End If
    Next comp
    Exit Sub
Skip:
    Call NavNotify("Thesis Library built, but the sheet event code could not be written (enable Trust access to the VBA project object model, or paste RR4/SheetThesis_Code.txt)", True)
End Sub
