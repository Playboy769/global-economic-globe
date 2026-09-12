Attribute VB_Name = "modNav"
Option Explicit

' ================================================================
'  RR4 NAV BAR v1 (2026-09-12) - typed page codes and commands
' ----------------------------------------------------------------
'  Rows 1-3 of every REPORT page (one column right on RR4, see NavLeft):
'    row 1   page code badge | command cell (grey = input)
'            (v4.4: no page title / status text on the bar any more -
'            NavNotify messages go to the Excel status bar only)
'    row 2   page codes   (the current page in white)
'    row 3   action codes
'  Type a code into that grey cell and press Enter; Workbook_SheetChange
'  hands it to RunNavCommand, which clears the cell again.
'
'  Pages   P RR4 . R Realized . T Transactions . H HistoryLog
'          V Volatility180D . VT Tickers Volatility
'          (D DrawdownChart removed 2026-09-12 with DrawDownShadow.bas)
'          C HoldingsCorr . CC Correlation . CR Company research
'  Actions UP update dashboard . ADD / DEL trade forms . V! D! C! recalc
'          and show that page . DBG system debug . CLEARALL wipe all data
'          (ClearAllData keeps its own Yes/No confirmation)
'  Jumping to a page never recalculates it - the "!" codes do that.
'
'  Which pages carry the bar:
'    RR4 reserves rows 1-3 in its own layout (PortfolioDashboard_v3,
'    RR4_TOP) and draws the bar itself.
'    V / VT / D / C / CC are drawn from row 1 by their own routines,
'    so those routines call NavStrip first (delete the 3 rows when they
'    are there) and NavAdd last (insert 3 rows, draw the bar). A hidden
'    sheet-level name RR4NAV marks a page that currently carries the rows,
'    which is what stops a redraw from stacking a second set.
'    VT and CC keep a typed input in B2; their sheet code finds it with
'    NavOffset (B5 while the bar is there). Those pages are not column-
'    shifted - only RR4 has the blank column A.
'    Data pages (R / T / H / CR) get no bar on purpose: many modules read
'    them by fixed row (header row 1, data from row 2).
' ================================================================
Public Const NAV_ROWS     As Long = 3
Private Const NAV_MARK    As String = "RR4NAV"

' Whether the last NavNotify was an error (NavEcho repaints it in red).
' Module-level declarations must precede every procedure in VBA - putting
' this next to NavNotify further down is a compile error for the whole
' module ("only comments may appear after End Sub").
Private m_lastIsErr As Boolean

' The bar starts in column A on every page except RR4, whose own layout keeps
' column A as a blank spacer (RR4_LEFT), so there it starts in B.
Private Function NavLeft(ByVal ws As Worksheet) As Long
    If NavPageCode(ws) = "P" Then NavLeft = RR4_LEFT
End Function

' Address of the page's command cell - B1, or C1 on RR4.
Public Function NavCmdCell(ByVal ws As Object) As String
    If Not TypeOf ws Is Worksheet Then Exit Function
    NavCmdCell = ws.cells(1, 2 + NavLeft(ws)).Address(False, False)
End Function

' --- code -> sheet tab name ---------------------------------------
Public Function NavSheetName(ByVal code As String) As String
    Select Case UCase(code)
        Case "P":  NavSheetName = "RR4"
        Case "R":  NavSheetName = "Realized"
        Case "T":  NavSheetName = "Transactions"
        Case "H":  NavSheetName = "HistoryLog"
        Case "V":  NavSheetName = "Volatility180D"
        Case "VT": NavSheetName = "Tickers Volatility"
        Case "C":  NavSheetName = "HoldingsCorr"
        Case "CC": NavSheetName = "Correlation"
        Case "RG": NavSheetName = "RRG"
        Case "CR": NavSheetName = "Company research"
    End Select
End Function

' Code of a report page (one that carries the bar), "" for anything else.
Public Function NavPageCode(ByVal ws As Object) As String
    If Not TypeOf ws Is Worksheet Then Exit Function
    Dim c As Variant
    For Each c In Array("P", "V", "VT", "C", "CC", "RG")
        If StrComp(ws.Name, NavSheetName(CStr(c)), vbTextCompare) = 0 Then
            NavPageCode = CStr(c)
            Exit Function
        End If
    Next c
End Function

Public Function NavHasRows(ByVal ws As Worksheet) As Boolean
    If NavPageCode(ws) = "P" Then
        NavHasRows = True           ' RR4 layout reserves rows 1-3
        Exit Function
    End If
    Dim nm As Name
    For Each nm In ws.Names
        If Right(nm.Name, Len(NAV_MARK) + 1) = "!" & NAV_MARK Then
            NavHasRows = True
            Exit Function
        End If
    Next nm
End Function

' Rows the bar currently pushes a page's own layout down by.
Public Function NavOffset(ByVal ws As Worksheet) As Long
    If NavPageCode(ws) = "P" Then Exit Function     ' RR4 layout already counts them
    If NavHasRows(ws) Then NavOffset = NAV_ROWS
End Function

' Take the bar rows off so a routine can redraw its page from row 1.
Public Sub NavStrip(ByVal ws As Worksheet)
    If NavPageCode(ws) = "P" Then Exit Sub
    If Not NavHasRows(ws) Then Exit Sub
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fin
    Call ShapesMoveOnly(ws)         ' see ShapesMoveOnly: charts anchored in rows 1-3
    ws.Rows("1:" & NAV_ROWS).Delete
    Dim k As Long
    For k = ws.Names.count To 1 Step -1
        If Right(ws.Names(k).Name, Len(NAV_MARK) + 1) = "!" & NAV_MARK Then ws.Names(k).Delete
    Next k
Fin:
    Application.EnableEvents = prevEv
End Sub

' Shapes default to xlMoveAndSize, which RESIZES them when rows inside their
' span are inserted or deleted - a chart anchored at row 1 would shrink on
' every strip/add cycle. xlMove keeps the size and just slides the shape
' with the rows, which is what the bar needs.
Public Sub ShapesMoveOnly(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long
    For i = 1 To ws.Shapes.count
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
    If code <> "P" And Not NavHasRows(ws) Then
        Call ShapesMoveOnly(ws)
        ws.Rows("1:" & NAV_ROWS).Insert Shift:=xlDown
        ws.Names.Add Name:=NAV_MARK, RefersTo:="=TRUE", Visible:=False
    End If
    Call DrawNavRows(ws, code)
    ' VT / CC keep a typed input in B2 (B5 with the bar): paint it like every
    ' other input cell (RR4_INPUT_BG dark grey, white text)
    If code = "VT" Or code = "CC" Then
        With ws.Range("B2").Offset(NAV_ROWS, 0)
            .Interior.Color = RR4_INPUT_BG
            .Font.Color = RR4_INPUT_FG
            .Font.Bold = True
        End With
    End If
Fin:
    Application.EnableEvents = prevEv
End Sub

' ----------------------------------------------------------------
' Paint rows 1-3. On RR4 the bar owns A:S only - T1/T2 hold the inception /
' capital config and X1 the ticker-panel tracker, and column A stays blank.
' ----------------------------------------------------------------
Public Sub DrawNavRows(ByVal ws As Worksheet, ByVal code As String)
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fin

    Dim off As Long: off = NavLeft(ws)
    Dim area As Range
    If code = "P" Then
        ' stop short of T1/T2 (the inception / capital config cells)
        Set area = ws.Range(ws.cells(1, 1), ws.cells(3, 19))
    Else
        Set area = ws.Range("A1:Z3")
    End If
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
    ' rows 2/3 are taller than the text needs and vertically centred, so the
    ' three lines sit apart without extra rows (v4.3 - was 16 / 18)
    ws.Rows(1).RowHeight = 24
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 22

    ' row 1: badge | command cell | title
    With ws.cells(1, 1 + off)
        .Value = code
        .Interior.Color = RR4_ACCENT
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.cells(1, 2 + off)
        .NumberFormat = "@"
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Size = 10
        .Font.Bold = True
    End With

    ' row 2: pages
    Dim pages As Variant
    pages = Array("P", "PORTFOLIO", "R", "REALIZED", "T", "TRANS", "H", "HISTORY", _
                  "V", "VOL", "VT", "TKRVOL", _
                  "C", "HOLDCORR", "CC", "SECTORCORR", "RG", "RRG", "CR", "RESEARCH")
    Call WriteCodeLine(ws.cells(2, 1 + off), pages, code, RR4_ACCENT)

    ' row 3: actions
    Dim acts As Variant
    acts = Array("UP", "UPDATE", "ADD", "TRADE", "DEL", "DELETE", "V!", "RECALC VOL", _
                 "C!", "CORR", "RG!", "RRG", "DBG", "DEBUG", "CLEARALL", "WIPE ALL DATA")
    Call WriteCodeLine(ws.cells(3, 1 + off), acts, "", RGB(0, 200, 255))

    ' divider under the bar: dark grey, starting at the bar's first column
    ' (B on RR4, so the blank spacer column A carries no line)
    With ws.Range(ws.cells(3, 1 + off), area.cells(3, area.Columns.count)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

Fin:
    ' leaving EnableEvents off would kill every input cell in the workbook
    Application.EnableEvents = prevEv
End Sub

' "CODE LABEL   CODE LABEL ..." in one cell (it overflows to the right).
' v4.3: everything is dimmed - codes at half brightness of codeColor,
' labels dark grey - and only the current page's pair is lit (code in the
' full colour, label light grey). No underline.
Private Sub WriteCodeLine(ByVal cell As Range, pairs As Variant, ByVal current As String, _
                          ByVal codeColor As Long)
    Dim s As String, i As Long
    For i = LBound(pairs) To UBound(pairs) Step 2
        s = s & pairs(i) & " " & pairs(i + 1) & "    "
    Next i
    cell.NumberFormat = "@"
    cell.Value = s
    cell.Font.Color = RGB(120, 120, 120)        ' dimmed labels
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

' v4.4: the bar no longer carries a title / status line - everything goes
' to the Excel status bar (bottom of the window). Kept as the single place
' status text is routed through so the callers did not have to change.
Public Sub NavStatus(ByVal ws As Worksheet, ByVal msg As String, ByVal isErr As Boolean)
    m_lastIsErr = isErr
    If Len(msg) > 0 Then Application.StatusBar = IIf(isErr, "ERROR - ", "") & msg
End Sub

' ================================================================
'  NavNotify replaces the old "done" / "nothing to do" MsgBoxes of the
'  routines the nav commands run (UP, V!, D!, C!, DBG, CLEARALL, the VT
'  and CC inputs): the message goes to the Excel status bar and to the
'  status line of the page in front - no dialog to click away. Real
'  errors and the CLEARALL Yes/No confirmation still use MsgBox.
' ================================================================
Public Sub NavNotify(ByVal msg As String, Optional ByVal isErr As Boolean = False)
    m_lastIsErr = isErr
    Application.StatusBar = IIf(isErr, "ERROR - ", "") & msg
End Sub

' v4.4: nothing to repaint on the page any more; kept so the sheet code /
' RunNavCommand callers still compile.
Public Sub NavEcho()
End Sub

' ================================================================
'  Command dispatcher (called from ThisWorkbook.Workbook_SheetChange)
' ================================================================
Public Sub RunNavCommand(ByVal raw As String, ByVal src As Worksheet)
    Dim cmd As String: cmd = UCase(Trim(raw))

    ' empty the command cell so the next code can be typed straight in
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    src.cells(1, 2 + NavLeft(src)).Value = ""
    Application.EnableEvents = prevEv
    If cmd = "" Then Exit Sub

    Select Case cmd
        Case "P", "R", "T", "H", "V", "VT", "C", "CC", "RG", "CR"
            Call NavGoto(cmd, src)
            Exit Sub                ' a jump has no result to echo
        Case "UP"
            Call RebuildPortfolioDashboard
        Case "ADD"
            Call OpenTransactionForm
        Case "DEL"
            Call OpenDeleteForm
        Case "V!"
            Call UpdatePortfolioVolatility
            Call NavGoto("V", src)
        Case "C!"
            Call BuildHoldingsCorrelation
        Case "RG!"
            Call BuildRRG
        Case "DBG"
            Call RunSystemDebug
        Case "CLEARALL"
            Call ClearAllData
        Case Else
            Call NavStatus(src, "UNKNOWN CODE [" & cmd & "]", True)
            Exit Sub
    End Select
    Call NavEcho
End Sub

' Jump to a page. A report page built before the bar existed gets it
' here, the first time it is visited.
Public Sub NavGoto(ByVal code As String, ByVal src As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NavSheetName(code))
    On Error GoTo 0
    If ws Is Nothing Then
        Call NavStatus(src, "[" & code & "] " & NavSheetName(code) & " is not built yet" & _
                       IIf(code = "V" Or code = "C" Or code = "RG", " - run " & code & "!", ""), True)
        Exit Sub
    End If
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    If NavPageCode(ws) <> "" And NavPageCode(ws) <> "P" And Not NavHasRows(ws) Then
        Call NavAdd(ws, code)
    End If
    ws.Activate
    If NavPageCode(ws) <> "" Then
        ws.cells(1, 2 + NavLeft(ws)).Select
    Else
        ws.Range("A1").Select
    End If
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub
