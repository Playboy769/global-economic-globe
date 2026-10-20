Attribute VB_Name = "modNav"
Option Explicit

' ================================================================
'  RR5 NAV BAR v1 (2026-10-20) - typed page codes, adapted from RR4's
'  modNav.bas (see Claudecode/RR4/modNav.bas for the original and its
'  full design notes). RR5 only has two pages so this is a trimmed copy:
'  no page carries a "baked-in" layout the way RR4's own P page does -
'  every page here is a normal report page that gets the bar inserted.
' ----------------------------------------------------------------
'  Rows 1-3 of every report page (one column right of its own content,
'  see NavLeft):
'    row 1   page code badge | command cell (grey = input)
'    row 2   page codes
'    row 3   action codes
'  Type a code into the grey cell and press Enter; Workbook_SheetChange
'  (ThisWorkbook) hands it to RunNavCommand, which clears the cell again.
'
'  Pages   R5 RR5 (main dashboard) . T Transactions
'  Actions UP RebuildPositionsFromTransactions (rebuild + fetch prices +
'             realize closed trades + daily snapshot - RR5_Core's own
'             "do everything" entry point)
'          DBG none yet
'  Jumping to a page never recalculates it - "UP" does that.
'
'  Every report page reserves row 1 + column A blank; the bar sits in
'  rows 2-4 from column B. NavAdd inserts the blank row + column together
'  with the 3 bar rows and records the geometry in the RR5NAV hidden
'  sheet-level name ("top|left"); NavStrip removes exactly what the mark
'  says. Post-build readers (RR5_Core / RR5_Dashboard) add NavOffset to
'  their row numbers and NavLeft to their column numbers wherever they
'  touch the RR5 or Transactions sheet - see the "NAV OFFSET" comments
'  sprinkled through those two modules.
' ================================================================
Public Const NAV_ROWS As Long = 3
Private Const NAV_MARK As String = "RR5NAV"

Private m_lastIsErr As Boolean

' Colors - lifted verbatim from RR4's PortfolioDashboard_v3.bas (RR4_ACCENT /
' RR4_INPUT_BG / RR4_INPUT_FG / RR4_LINE) per the user's choice to reuse the
' same visual convention here.
Public Const RR5_ACCENT   As Long = 25800      ' RGB(200,100,0)
Public Const RR5_INPUT_BG As Long = 2631720    ' RGB(40,40,40)
Public Const RR5_INPUT_FG As Long = 16777215   ' RGB(255,255,255) white
Public Const RR5_LINE     As Long = 4605510    ' RGB(70,70,70)

' --- code -> sheet tab name ---------------------------------------
Public Function NavSheetName(ByVal code As String) As String
    Select Case UCase(code)
        Case "R5": NavSheetName = "RR5"
        Case "T":  NavSheetName = "Transactions"
    End Select
End Function

' Code of a report page (one that carries the bar), "" for anything else.
Public Function NavPageCode(ByVal ws As Object) As String
    If Not TypeOf ws Is Worksheet Then Exit Function
    Dim c As Variant
    For Each c In Array("R5", "T")
        If StrComp(ws.Name, NavSheetName(CStr(c)), vbTextCompare) = 0 Then
            NavPageCode = CStr(c)
            Exit Function
        End If
    Next c
End Function

' --- geometry: every report page here gets the same blank row+col ---
Private Function NavWantTop(ByVal ws As Worksheet) As Long
    NavWantTop = 1
End Function

Private Function NavWantLeft(ByVal ws As Worksheet) As Long
    NavWantLeft = 1
End Function

' Geometry currently ON the sheet, from the mark (0/0 when no bar rows).
Private Sub NavGeom(ByVal ws As Worksheet, ByRef top As Long, ByRef left As Long)
    top = 0: left = 0
    Dim nm As Name
    For Each nm In ws.Names
        If Right(nm.Name, Len(NAV_MARK) + 1) = "!" & NAV_MARK Then
            Dim v As String: v = Replace(Replace(nm.RefersTo, "=", ""), """", "")
            If InStr(v, "|") > 0 Then
                top = CLng(Val(Split(v, "|")(0))): left = CLng(Val(Split(v, "|")(1)))
            End If
            Exit Sub
        End If
    Next nm
End Sub

' Columns the page content sits right of (blank column A) - PUBLIC so
' RR5_Core / RR5_Dashboard can offset their post-build cell reads.
Public Function NavLeft(ByVal ws As Worksheet) As Long
    Dim t As Long, l As Long
    Call NavGeom(ws, t, l)
    NavLeft = l
End Function

' Blank rows above the bar on this sheet right now.
Private Function NavTop(ByVal ws As Worksheet) As Long
    Dim t As Long, l As Long
    Call NavGeom(ws, t, l)
    NavTop = t
End Function

' Rows the bar block takes in total (blank rows above + the 3 bar rows).
Private Function NavBlock(ByVal ws As Worksheet) As Long
    NavBlock = NAV_ROWS + NavTop(ws)
End Function

' Rows the bar currently pushes a page's own layout down by - PUBLIC so
' RR5_Core / RR5_Dashboard can offset their row numbers.
Public Function NavOffset(ByVal ws As Worksheet) As Long
    If NavHasRows(ws) Then NavOffset = NavBlock(ws)
End Function

Public Function NavHasRows(ByVal ws As Worksheet) As Boolean
    Dim nm As Name
    For Each nm In ws.Names
        If Right(nm.Name, Len(NAV_MARK) + 1) = "!" & NAV_MARK Then
            NavHasRows = True
            Exit Function
        End If
    Next nm
End Function

' Address of the page's command cell - "C2" once the bar is on.
Public Function NavCmdCell(ByVal ws As Object) As String
    If Not TypeOf ws Is Worksheet Then Exit Function
    NavCmdCell = ws.Cells(1 + NavTop(ws), 2 + NavLeft(ws)).Address(False, False)
End Function

' Take the bar rows off so a routine can redraw its page from row 1.
Public Sub NavStrip(ByVal ws As Worksheet)
    If Not NavHasRows(ws) Then Exit Sub
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fin
    Call ShapesMoveOnly(ws)
    Dim gt As Long, gl As Long
    Call NavGeom(ws, gt, gl)
    ws.Rows("1:" & (NAV_ROWS + gt)).Delete
    If gl > 0 Then ws.Range(ws.Columns(1), ws.Columns(gl)).Delete
    Dim k As Long
    For k = ws.Names.Count To 1 Step -1
        If Right(ws.Names(k).Name, Len(NAV_MARK) + 1) = "!" & NAV_MARK Then ws.Names(k).Delete
    Next k
Fin:
    Application.EnableEvents = prevEv
End Sub

' Shapes default to xlMoveAndSize, which RESIZES them when rows/cols
' inside their span are inserted or deleted. xlMove keeps the size and
' just slides the shape with the rows, which is what the bar needs.
Public Sub ShapesMoveOnly(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long
    For i = 1 To ws.Shapes.Count
        ws.Shapes(i).Placement = xlMove
    Next i
    On Error GoTo 0
End Sub

' Put the bar rows back on top of a page (insert once, redraw always).
Public Sub NavAdd(ByVal ws As Worksheet, Optional ByVal code As String = "")
    If code = "" Then code = NavPageCode(ws)
    If code = "" Then Exit Sub
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fin
    If Not NavHasRows(ws) Then
        Call ShapesMoveOnly(ws)
        Dim wt As Long, wl As Long
        wt = NavWantTop(ws): wl = NavWantLeft(ws)
        ws.Rows("1:" & (NAV_ROWS + wt)).Insert Shift:=xlDown
        If wl > 0 Then
            ws.Range(ws.Columns(1), ws.Columns(wl)).Insert Shift:=xlToRight
            Dim ci As Long
            For ci = 1 To wl: ws.Columns(ci).ColumnWidth = 3: Next ci
        End If
        ws.Names.Add Name:=NAV_MARK, RefersTo:="=""" & wt & "|" & wl & """", Visible:=False
    End If
    Call DrawNavRows(ws, code)
Fin:
    Application.EnableEvents = prevEv
End Sub

' ----------------------------------------------------------------
' Paint rows 1-3.
' ----------------------------------------------------------------
Public Sub DrawNavRows(ByVal ws As Worksheet, ByVal code As String)
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fin

    Dim off As Long: off = NavLeft(ws)
    Dim top As Long: top = NavTop(ws)
    Dim area As Range
    Set area = ws.Range(ws.Cells(1 + top, 1), ws.Cells(3 + top, 26))
    With area
        .Clear
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(150, 150, 150)
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
    If top > 0 Then ws.Rows(top).RowHeight = 14      ' the blank row above the bar
    ws.Rows(1 + top).RowHeight = 24
    ws.Rows(2 + top).RowHeight = 22
    ws.Rows(3 + top).RowHeight = 22

    ' row 1: badge | command cell
    With ws.Cells(1 + top, 1 + off)
        .Value = code
        .Interior.Color = RR5_ACCENT
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(1 + top, 2 + off)
        .NumberFormat = "@"
        .Interior.Color = RR5_INPUT_BG
        .Font.Color = RR5_INPUT_FG
        .Font.Size = 10
        .Font.Bold = True
    End With

    ' row 2: pages
    Dim pages As Variant
    pages = Array("R5", "RR5", "T", "TRANSACTION")
    Call WriteCodeLine(ws.Cells(2 + top, 1 + off), pages, code, RR5_ACCENT)

    ' row 3: actions
    Dim acts As Variant
    acts = Array("UP", "UPDATE")
    Call WriteCodeLine(ws.Cells(3 + top, 1 + off), acts, "", RGB(0, 200, 255))

    With ws.Range(ws.Cells(3 + top, 1 + off), area.Cells(3, area.Columns.Count)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR5_LINE
        .Weight = xlThin
    End With

Fin:
    Application.EnableEvents = prevEv
End Sub

' "CODE LABEL   CODE LABEL ..." in one cell (it overflows to the right).
Private Sub WriteCodeLine(ByVal cell As Range, pairs As Variant, ByVal current As String, _
                          ByVal codeColor As Long)
    Dim s As String, i As Long
    For i = LBound(pairs) To UBound(pairs) Step 2
        s = s & pairs(i) & " " & pairs(i + 1) & "    "
    Next i
    cell.NumberFormat = "@"
    cell.Value = s
    cell.Font.Color = RGB(120, 120, 120)
    Dim dimCode As Long: dimCode = DimColor(codeColor, 0.6)
    Dim pos As Long: pos = 1
    For i = LBound(pairs) To UBound(pairs) Step 2
        Dim cLen As Long: cLen = Len(pairs(i))
        Dim lLen As Long: lLen = Len(pairs(i + 1))
        If StrComp(pairs(i), current, vbTextCompare) = 0 Then
            With cell.Characters(pos, cLen).Font
                .Color = codeColor
                .Bold = True
            End With
            With cell.Characters(pos + cLen + 1, lLen).Font
                .Color = RGB(235, 235, 235)
                .Bold = True
            End With
        Else
            With cell.Characters(pos, cLen).Font
                .Color = dimCode
                .Bold = True
            End With
        End If
        pos = pos + cLen + 1 + lLen + 4
    Next i
End Sub

' Scale an RGB Long towards black (factor 1 = unchanged, 0 = black).
Private Function DimColor(ByVal c As Long, ByVal factor As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = c Mod 256: g = (c \ 256) Mod 256: b = (c \ 65536) Mod 256
    DimColor = RGB(CLng(r * factor), CLng(g * factor), CLng(b * factor))
End Function

Public Sub NavStatus(ByVal ws As Worksheet, ByVal msg As String, ByVal isErr As Boolean)
    m_lastIsErr = isErr
    If Len(msg) > 0 Then Application.StatusBar = IIf(isErr, "ERROR - ", "") & msg
End Sub

Public Sub NavNotify(ByVal msg As String, Optional ByVal isErr As Boolean = False)
    m_lastIsErr = isErr
    Application.StatusBar = IIf(isErr, "ERROR - ", "") & msg
End Sub

' ================================================================
'  Command dispatcher (called from ThisWorkbook.Workbook_SheetChange)
' ================================================================
Public Sub RunNavCommand(ByVal raw As String, ByVal src As Worksheet)
    Dim cmd As String: cmd = UCase(Trim(raw))

    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    src.Cells(1 + NavTop(src), 2 + NavLeft(src)).Value = ""
    Application.EnableEvents = prevEv
    If cmd = "" Then Exit Sub

    Select Case cmd
        Case "R5", "T"
            Call NavGoto(cmd, src)
            Exit Sub                ' a jump has no result to echo
        Case "UP"
            Call RR5_Core.RebuildPositionsFromTransactions
        Case Else
            Call NavStatus(src, "UNKNOWN CODE [" & cmd & "]", True)
            Exit Sub
    End Select
    Call NavNotify("UP done " & Format(Now, "hh:mm:ss"))
End Sub

' Jump to a page. A report page built before the bar existed gets it
' here, the first time it is visited.
Public Sub NavGoto(ByVal code As String, ByVal src As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NavSheetName(code))
    On Error GoTo 0
    If ws Is Nothing Then
        Call NavStatus(src, "[" & code & "] " & NavSheetName(code) & " not found", True)
        Exit Sub
    End If
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    If NavPageCode(ws) <> "" And Not NavHasRows(ws) Then
        Call NavAdd(ws, code)
    End If
    ws.Activate
    If NavPageCode(ws) <> "" Then
        ws.Cells(1 + NavTop(ws), 2 + NavLeft(ws)).Select
    Else
        ws.Range("A1").Select
    End If
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub
