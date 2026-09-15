Attribute VB_Name = "modValuationPage"
Option Explicit

' ================================================================
'  VALUATION (nav code VA, VA! = refetch) - sheet "Valuation"   2026-09-15
' ----------------------------------------------------------------
'  ROIC-based relative valuation, no WACC.  All numbers come from
'  RR4/valuation.py (SEC companyfacts / MOPS iXBRL / Yahoo prices): this
'  module writes a request file, shells python, reads the tab-separated
'  result back and only draws the page.  Definitions match
'  SEC-Filing-Fetcher/modValuation.bas (IC = assets - current liabilities,
'  NOPAT = EBIT x (1 - tax)); see the script header for the rest.
'
'  Peer group = the tblGroups group that contains the ticker (Sanner.
'  GetGroupNames / GetSectorTickers): theme groups first (>= 4 members),
'  then the numbered sector groups, TW before US; the GROUP input can
'  override it, and with no group at all only the ROIC side is filled.  Page layout (page rows; the bar adds a blank row, 3 bar
'  rows and a blank column A - read inputs through NavOffset / NavLeft):
'    row 1     VALUATION title | context
'    row 2-3   TICKER <GO> (A) . MKT (C) . GROUP (E)   label above input
'    row 5-    left  A:C  ROIC summary        right H:J  relative + fair value
'    then      QUARTERLY HISTORY (8 quarters), PEER GROUP table,
'              VA_SCATTER chart (EV/EBIT vs ROIC) to the right of the tables,
'              THESIS lines ([4]-[6] of the 500-word thesis), LOG.
'  Event code (Worksheet_Change -> ValuationChange) is written by
'  EnsureSheetCode; RR4/SheetValuation_Code.txt is the record.
'  Pure ASCII (VBE import rule): the script path's Chinese folder name is
'  built with ChrW.
' ================================================================

Public Const VAL_SHEET As String = "Valuation"
Private Const CHART_NAME As String = "VA_SCATTER"

Private Const PG_TITLE As Long = 1
Private Const PG_LBL As Long = 2
Private Const PG_IN As Long = 3
Private Const COL_TK As Long = 1
Private Const COL_MKT As Long = 3
Private Const COL_GRP As Long = 5
Private Const PG_BODY As Long = 5
Private Const RIGHT_COL As Long = 8                  ' H: relative / fair blocks
Private Const LAST_CLEAR_COL As Long = 40

Private Const FONT_FACE As String = "Consolas"
Private Const CLR_TEXT As Long = 14540253            ' RGB(221,221,221)
Private Const CLR_MUTED As Long = 8553090            ' RGB(130,130,130)
Private Const CLR_BANNER As Long = 1842204           ' RGB(28,28,28)
Private Const CLR_GOOD As Long = 5287936             ' RGB(0,176,80)  (green = above hurdle)
Private Const CLR_BAD As Long = 3355647              ' RGB(255,50,50)
Private Const CLR_FLAG As Long = 49407               ' RGB(255,192,0)

' parsed result: block name -> Collection of Variant(0..n) cell arrays
Private m_blocks As Object

' ----------------------------------------------------------------
'  Entry points
' ----------------------------------------------------------------
Public Sub ShowValuation(ByVal rawTicker As String, Optional ByVal marketOverride As String = "", _
                         Optional ByVal groupOverride As String = "", Optional ByVal goPage As Boolean = False)
    Dim ws As Worksheet: Set ws = EnsureValuationSheet()
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fail

    Call NavStrip(ws)
    Dim tk As String: tk = UCase$(Trim$(rawTicker))
    Dim mkt As String: mkt = UCase$(Trim$(marketOverride))
    Dim grp As String: grp = Trim$(groupOverride)
    Call DrawShell(ws, tk, mkt, grp)
    If tk = "" Then GoTo Done

    Dim bare As String: bare = tk
    If Right$(bare, 4) = ".TWO" Then bare = Left$(bare, Len(bare) - 4): mkt = "TW"
    If Right$(bare, 3) = ".TW" Then bare = Left$(bare, Len(bare) - 3): mkt = "TW"

    ' peer group from tblGroups
    Dim gMkt As String, peers As Variant
    If grp = "" Then
        grp = FindGroupOf(bare, mkt, gMkt)
    Else
        gMkt = GroupMarketOf(grp)
    End If
    ' an all-digit code is a TW stock whatever the group table says
    If IsNumeric(bare) Then mkt = "TW"
    If mkt = "" Then mkt = IIf(gMkt <> "", gMkt, "US")
    If grp <> "" Then peers = GetSectorTickers(mkt, grp)
    ' no group = ROIC side only; the relative / fair-value blocks read n/a
    Dim nPeers As Long
    If IsEmpty(peers) Then
        peers = Array()
        grp = ""
    Else
        nPeers = UBound(peers) - LBound(peers)
    End If
    ws.cells(PG_IN, COL_GRP).Value = grp

    Call NavNotify("VALUATION fetching " & bare & IIf(nPeers > 0, " + " & nPeers & " peers", " (no peer group)") & " (" & mkt & ") ...")
    Dim outPath As String, errMsg As String
    If Not RunScript(bare, mkt, grp, peers, outPath, errMsg) Then
        Call WriteFlag(ws, PG_BODY, "valuation.py failed: " & errMsg)
        Call NavNotify("VALUATION error: " & errMsg, True)
        GoTo Done
    End If
    If Not LoadResult(outPath, errMsg) Then
        Call WriteFlag(ws, PG_BODY, "valuation.py: " & errMsg)
        Call NavNotify("VALUATION error: " & errMsg, True)
        GoTo Done
    End If

    Call RenderPage(ws)
    Call NavNotify("VALUATION " & bare & " done - " & Meta("peers_used") & " peers used, " & Meta("elapsed_s") & "s")

Done:
    Call FinishPage(ws)
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
    If goPage Then Call NavGoto("VA", ws)
    Exit Sub
Fail:
    Dim msg As String: msg = Err.Description
    Resume FailOut
FailOut:
    On Error Resume Next
    Call WriteFlag(ws, PG_BODY, "Error: " & msg)
    Call NavNotify("VALUATION error: " & msg, True)
    GoTo Done
End Sub

' VA! - refetch whatever the inputs hold.
Public Sub RefreshValuation()
    Dim ws As Worksheet: Set ws = EnsureValuationSheet()
    Dim r As Long: r = PG_IN + NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Call ShowValuation(CStr(ws.cells(r, COL_TK + lc).Value), CStr(ws.cells(r, COL_MKT + lc).Value), _
                       CStr(ws.cells(r, COL_GRP + lc).Value), True)
End Sub

' Sheet event (Worksheet_Change, written by EnsureSheetCode).
Public Sub ValuationChange(ws As Worksheet, ByVal Target As Range)
    If Target.CountLarge > 4 Then Exit Sub
    Dim r As Long: r = PG_IN + NavOffset(ws)
    Dim lc As Long: lc = NavLeft(ws)
    Dim inputs As Range
    Set inputs = Union(ws.cells(r, COL_TK + lc), ws.cells(r, COL_MKT + lc), ws.cells(r, COL_GRP + lc))
    If Intersect(Target, inputs) Is Nothing Then Exit Sub
    Call ShowValuation(CStr(ws.cells(r, COL_TK + lc).Value), CStr(ws.cells(r, COL_MKT + lc).Value), _
                       CStr(ws.cells(r, COL_GRP + lc).Value))
End Sub

' ----------------------------------------------------------------
'  Peer group lookup (tblGroups)
' ----------------------------------------------------------------
Private Function BareTicker(ByVal s As String) As String
    s = UCase$(Trim$(s))
    If Right$(s, 4) = ".TWO" Then s = Left$(s, Len(s) - 4)
    If Right$(s, 3) = ".TW" Then s = Left$(s, Len(s) - 3)
    BareTicker = s
End Function

' Group whose ticker list contains the ticker.  Theme groups (names that do
' not start with "NN. ") win over the numbered sector groups, and a theme
' group with fewer than 4 members is skipped so the multiple band has
' something to stand on.  TW is scanned before US when no market is given;
' gMkt returns the market it was found in.
Private Function FindGroupOf(ByVal bare As String, ByVal mkt As String, ByRef gMkt As String) As String
    Dim mkts As Variant
    If mkt = "" Then mkts = Array("TW", "US") Else mkts = Array(mkt)
    Dim pass As Long, m As Variant, g As Variant, t As Variant, names As Variant, tks As Variant
    For pass = 1 To 2                      ' 1 = theme groups, 2 = numbered groups
        For Each m In mkts
            Dim mm As String: mm = CStr(m)
            names = GetGroupNames(mm)
            If Not IsEmpty(names) Then
                For Each g In names
                    Dim gg As String: gg = CStr(g)
                    If IsNumberedGroup(gg) = (pass = 2) Then
                        tks = GetSectorTickers(mm, gg)
                        If Not IsEmpty(tks) Then
                            If pass = 2 Or UBound(tks) - LBound(tks) + 1 >= 4 Then
                                For Each t In tks
                                    If BareTicker(CStr(t)) = bare Then
                                        gMkt = mm
                                        FindGroupOf = gg
                                        Exit Function
                                    End If
                                Next t
                            End If
                        End If
                    End If
                Next g
            End If
        Next m
    Next pass
End Function

' "02. ..." style sector group (two digits, a dot, a space)
Private Function IsNumberedGroup(ByVal g As String) As Boolean
    g = Trim$(g)
    If Len(g) < 4 Then Exit Function
    IsNumberedGroup = (Mid$(g, 1, 1) Like "#" And Mid$(g, 2, 1) Like "#" And Mid$(g, 3, 2) = ". ")
End Function

Private Function GroupMarketOf(ByVal grp As String) As String
    Dim m As Variant, names As Variant, g As Variant
    For Each m In Array("TW", "US")
        Dim mm As String: mm = CStr(m)
        names = GetGroupNames(mm)
        If Not IsEmpty(names) Then
            For Each g In names
                If StrComp(Trim$(CStr(g)), grp, vbTextCompare) = 0 Then GroupMarketOf = mm: Exit Function
            Next g
        End If
    Next m
End Function

' ----------------------------------------------------------------
'  Python round trip
' ----------------------------------------------------------------
Private Function ScriptPath() As String
    ' ...\OneDrive\<desktop>\Claudecode\RR4\valuation.py
    ScriptPath = Environ("USERPROFILE") & "\OneDrive\" & ChrW(&H684C) & ChrW(&H9762) & "\Claudecode\RR4\valuation.py"
End Function

Private Function WorkDir() As String
    Dim d As String: d = Environ("TEMP") & "\rr4-valuation"
    If Dir(d, vbDirectory) = "" Then MkDir d
    WorkDir = d
End Function

Private Function JsonStr(ByVal s As String) As String
    ' ASCII-safe JSON string (non-ASCII -> \uXXXX so the file stays ANSI)
    Dim i As Long, c As Long, o As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        Select Case c
            Case 34: o = o & "\"""
            Case 92: o = o & "\\"
            Case Is < 32: o = o & "\u" & Right$("000" & Hex(c), 4)
            Case Is > 126: o = o & "\u" & Right$("000" & Hex(c), 4)
            Case Else: o = o & ChrW(c)
        End Select
    Next i
    JsonStr = """" & o & """"
End Function

Private Function RunScript(ByVal tk As String, ByVal mkt As String, ByVal grp As String, peers As Variant, _
                           ByRef outPath As String, ByRef errMsg As String) As Boolean
    Dim sp As String: sp = ScriptPath()
    If Dir(sp) = "" Then errMsg = "script not found: " & sp: Exit Function
    Dim wd As String: wd = WorkDir()
    Dim reqPath As String: reqPath = wd & "\request.json"
    Dim errPath As String: errPath = wd & "\stderr.txt"
    outPath = wd & "\result.txt"

    Dim j As String, p As Variant, first As Boolean: first = True
    j = "{""ticker"": " & JsonStr(tk) & ", ""market"": " & JsonStr(mkt) & ", ""group"": " & JsonStr(grp) & ", ""peers"": ["
    If Not IsEmpty(peers) Then
        For Each p In peers
            If Not first Then j = j & ", "
            j = j & JsonStr(CStr(p))
            first = False
        Next p
    End If
    j = j & "]}"
    Dim fn As Integer: fn = FreeFile
    Open reqPath For Output As #fn
    Print #fn, j
    Close #fn
    On Error Resume Next
    Kill outPath: Kill errPath
    On Error GoTo 0

    ' hidden console, wait for exit; stderr captured for the error message
    Dim cmd As String
    cmd = "cmd /c ""python """ & sp & """ --req """ & reqPath & """ --out """ & outPath & """ 2> """ & errPath & """"""
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = sh.Run(cmd, 0, True)
    If Dir(outPath) = "" Then
        errMsg = "no result file (exit " & rc & "). " & Left$(ReadUtf8(errPath), 300)
        Exit Function
    End If
    RunScript = True
End Function

Private Function ReadUtf8(ByVal path As String) As String
    On Error GoTo Fin
    If Dir(path) = "" Then Exit Function
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.LoadFromFile path
    ReadUtf8 = stm.ReadText
    stm.Close
    If Left$(ReadUtf8, 1) = ChrW(&HFEFF) Then ReadUtf8 = Mid$(ReadUtf8, 2)
Fin:
End Function

' "#NAME" lines start a block; every other line is one tab-split row.
Private Function LoadResult(ByVal path As String, ByRef errMsg As String) As Boolean
    Set m_blocks = CreateObject("Scripting.Dictionary")
    Dim txt As String: txt = Replace(ReadUtf8(path), vbCr, "")
    Dim ln As Variant, cur As String
    For Each ln In Split(txt, vbLf)
        Dim s As String: s = CStr(ln)
        If Left$(s, 1) = "#" Then
            cur = Mid$(s, 2)
            If Not m_blocks.Exists(cur) Then m_blocks.Add cur, New Collection
        ElseIf cur <> "" And Len(s) > 0 Then
            m_blocks(cur).Add Split(s, vbTab)
        End If
    Next ln
    If m_blocks.Exists("ERROR") Then
        If m_blocks("ERROR").count > 0 Then errMsg = m_blocks("ERROR")(1)(0) Else errMsg = "unknown error"
        Exit Function
    End If
    If Not m_blocks.Exists("META") Then errMsg = "result file has no META block": Exit Function
    LoadResult = True
End Function

Private Function Meta(ByVal key As String) As String
    If m_blocks Is Nothing Then Exit Function
    If Not m_blocks.Exists("META") Then Exit Function
    Dim r As Variant
    For Each r In m_blocks("META")
        If r(0) = key Then
            If UBound(r) >= 1 Then Meta = r(1)
            Exit Function
        End If
    Next r
End Function

Private Function Cell(r As Variant, ByVal i As Long) As String
    If i <= UBound(r) Then Cell = r(i)
End Function

Private Function ToNum(ByVal s As String) As Variant
    If Len(s) = 0 Then ToNum = "" Else If IsNumeric(s) Then ToNum = CDbl(s) Else ToNum = s
End Function

' ----------------------------------------------------------------
'  Sheet / shell
' ----------------------------------------------------------------
Private Function EnsureValuationSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VAL_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Dim prev As Object: Set prev = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = VAL_SHEET
        ActiveWindow.DisplayGridlines = False
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If
    Set EnsureValuationSheet = ws
End Function

Private Sub DrawShell(ws As Worksheet, ByVal tk As String, ByVal mkt As String, ByVal grp As String)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
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
        .Value = "VALUATION"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 14
    End With
    ws.Rows(PG_TITLE).RowHeight = 24
    Call StackLabel(ws.cells(PG_LBL, COL_TK), "TICKER <GO>")
    Call StackLabel(ws.cells(PG_LBL, COL_MKT), "MKT")
    Call StackLabel(ws.cells(PG_LBL, COL_GRP), "GROUP")
    ws.Rows(PG_LBL).RowHeight = 18
    ws.Rows(PG_IN).RowHeight = 20
    Call InputCell(ws.cells(PG_IN, COL_TK), tk)
    Call InputCell(ws.cells(PG_IN, COL_MKT), mkt)
    Call InputCell(ws.cells(PG_IN, COL_GRP), grp)
    With ws.cells(PG_IN, COL_GRP + 2)
        .Value = "US ticker or TW code + Enter.  MKT blank = auto.  GROUP blank = the tblGroups group holding the ticker.  ROIC only, no WACC."
        .Font.Color = CLR_MUTED
    End With
    ws.Columns(1).ColumnWidth = 24
    ws.Columns(2).ColumnWidth = 14
    ws.Columns(3).ColumnWidth = 14
    For i = 4 To LAST_CLEAR_COL: ws.Columns(i).ColumnWidth = 12: Next i
    ws.Columns(RIGHT_COL).ColumnWidth = 24
End Sub

Private Sub StackLabel(c As Range, ByVal txt As String)
    With c
        .Value = txt
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
End Sub

Private Sub InputCell(c As Range, ByVal v As String)
    With c
        .NumberFormat = "@"
        .Value = v
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub FinishPage(ws As Worksheet)
    On Error Resume Next
    Call NavAdd(ws, "VA")
    Call EnsureSheetCode(ws)
    On Error GoTo 0
End Sub

Private Sub WriteFlag(ws As Worksheet, ByVal r As Long, ByVal txt As String)
    With ws.cells(r, 1)
        .Value = txt
        .Font.Color = CLR_FLAG
    End With
End Sub

Private Sub Section(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal txt As String, ByVal span As Long)
    With ws.Range(ws.cells(r, c), ws.cells(r, c + span - 1))
        .Interior.Color = CLR_BANNER
    End With
    With ws.cells(r, c)
        .Value = txt
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
    End With
    With ws.Range(ws.cells(r, c), ws.cells(r, c + span - 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With
End Sub

Private Sub HeaderRow(ws As Worksheet, ByVal r As Long, ByVal c As Long, hdr As Variant)
    Dim i As Long
    For i = 0 To UBound(hdr)
        With ws.cells(r, c + i)
            .Value = hdr(i)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .HorizontalAlignment = IIf(i = 0, xlLeft, xlRight)
        End With
    Next i
    With ws.Range(ws.cells(r, c), ws.cells(r, c + UBound(hdr))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RR4_LINE: .Weight = xlThin
    End With
End Sub

' Writes one "label / value / note" row with the value formatted by kind.
Private Sub KvRow(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal label As String, ByVal raw As String, _
                  ByVal kind As String, ByVal note As String)
    ws.cells(r, c).Value = label
    Dim v As Variant: v = ToNum(raw)
    With ws.cells(r, c + 1)
        .HorizontalAlignment = xlRight
        If kind = "txt" Or Not IsNumeric(raw) Then
            .NumberFormat = "@"
            .Value = IIf(Len(raw) = 0, "n/a", raw)
            .HorizontalAlignment = IIf(kind = "txt", xlLeft, xlRight)
            If Len(raw) = 0 Then .Font.Color = CLR_MUTED
        Else
            .Value = v
            Select Case kind
                Case "pct": .NumberFormat = "0.0%"
                Case "x": .NumberFormat = "0.0""x"""
                Case "px": .NumberFormat = "#,##0.00"
                Case Else: .NumberFormat = "#,##0,,;(#,##0,,);0"    ' millions
            End Select
        End If
        .Font.Bold = True
    End With
    With ws.cells(r, c + 2)
        .Value = note
        .Font.Color = CLR_MUTED
    End With
End Sub

' ----------------------------------------------------------------
'  Render
' ----------------------------------------------------------------
Private Sub RenderPage(ws As Worksheet)
    Dim unitNote As String: unitNote = Meta("unit")
    With ws.cells(PG_TITLE, 3)
        .Value = Meta("name") & "  (" & Meta("ticker") & ")  .  " & Meta("market") & "  .  group " & Meta("group") & _
                 "  .  " & Meta("src") & "  .  built " & Meta("built")
        .Font.Color = CLR_MUTED
    End With

    ' ---- left: ROIC summary
    Dim r As Long: r = PG_BODY
    Call Section(ws, r, 1, "ROIC  (" & unitNote & ", amounts in millions)", 3)
    r = r + 1
    Dim row As Variant
    Dim floorV As Double: floorV = Val(Meta("floor"))
    For Each row In m_blocks("SUMMARY")
        Call KvRow(ws, r, 1, Cell(row, 0), Cell(row, 1), Cell(row, 2), Cell(row, 3))
        ' hurdle colouring on the three ROIC readings
        If Left$(Cell(row, 0), 4) = "ROIC" And IsNumeric(Cell(row, 1)) Then
            ws.cells(r, 2).Font.Color = IIf(CDbl(Cell(row, 1)) >= floorV, CLR_GOOD, CLR_BAD)
        End If
        If Cell(row, 0) = "Last 4Q avg" Then
            Select Case Cell(row, 3)
                Case "IMPROVING": ws.cells(r, 3).Font.Color = CLR_GOOD
                Case "DETERIORATING": ws.cells(r, 3).Font.Color = CLR_BAD
            End Select
        End If
        r = r + 1
    Next row
    Dim leftEnd As Long: leftEnd = r

    ' ---- right: relative + fair value
    r = PG_BODY
    Call Section(ws, r, RIGHT_COL, "RELATIVE TO PEER GROUP", 3)
    r = r + 1
    For Each row In m_blocks("RELATIVE")
        Call KvRow(ws, r, RIGHT_COL, Cell(row, 0), Cell(row, 1), Cell(row, 2), Cell(row, 3))
        If Cell(row, 0) = "Premium / discount" And IsNumeric(Cell(row, 1)) Then
            ws.cells(r, RIGHT_COL + 1).Font.Color = IIf(CDbl(Cell(row, 1)) <= 0, CLR_GOOD, CLR_BAD)
        End If
        r = r + 1
    Next row
    r = r + 1
    Call Section(ws, r, RIGHT_COL, "FAIR VALUE BAND  (peer EV/EBIT x EBIT TTM - net debt) / shares", 3)
    r = r + 1
    For Each row In m_blocks("FAIR")
        Call KvRow(ws, r, RIGHT_COL, Cell(row, 0), Cell(row, 1), Cell(row, 2), Cell(row, 3))
        If Cell(row, 0) = "Safety margin" And IsNumeric(Cell(row, 1)) Then
            ws.cells(r, RIGHT_COL + 1).Font.Color = IIf(CDbl(Cell(row, 1)) >= 0.25, CLR_GOOD, CLR_BAD)
        End If
        If Cell(row, 0) = "Fair price (chosen)" Then ws.cells(r, RIGHT_COL + 1).Font.Color = RR4_ACCENT
        r = r + 1
    Next row
    If r > leftEnd Then leftEnd = r

    ' ---- quarterly history
    r = leftEnd + 1
    Call Section(ws, r, 1, "QUARTERLY HISTORY  (ROIC q x4 = Excel column convention; TTM = rolling four quarters)", 10)
    r = r + 1
    Call HeaderRow(ws, r, 1, Array("PERIOD", "EBIT Q", "NOPAT Q", "IC", "ROIC Qx4", "ROIC TTM", "TTM EX-CASH", "TTM STD", "NOPAT MGN", "IC TURN"))
    r = r + 1
    Dim first As Boolean: first = True
    For Each row In m_blocks("HISTORY")
        If first Then
            first = False
        Else
            Call WriteVals(ws, r, 1, row, Array("@", "M", "M", "M", "pct", "pct", "pct", "pct", "pct", "x"))
            If IsNumeric(Cell(row, 5)) Then ws.cells(r, 6).Font.Color = IIf(CDbl(Cell(row, 5)) >= floorV, CLR_GOOD, CLR_BAD)
            r = r + 1
        End If
    Next row
    Call WriteNote(ws, r, "Q4 is FY minus the first three quarters (SEC) / FY minus Q3 YTD (MOPS).  Tax rate outside 0-40% falls back to 21% / 20%.")
    r = r + 2

    ' ---- peers
    Call Section(ws, r, 1, "PEER GROUP  " & IIf(Meta("group") = "", "(none - type a GROUP name to compare)", Meta("group")) & "  (tblGroups, " & Meta("peers_used") & " usable for the multiple band)", 13)
    r = r + 1
    Call HeaderRow(ws, r, 1, Array("TICKER", "NAME", "PERIOD", "ROIC", "EV/EBIT", "P/B", "NOPAT MGN", "IC TURN", "MCAP", "EV", "EBIT TTM", "IC", "NOTE"))
    r = r + 1
    Dim peerFirst As Long: peerFirst = r
    first = True
    For Each row In m_blocks("PEERS")
        If first Then
            first = False
        Else
            Call WriteVals(ws, r, 1, row, Array("@", "@", "@", "pct", "x", "x", "pct", "x", "M", "M", "M", "M", "@"))
            ws.cells(r, 13).Font.Color = CLR_MUTED
            ws.cells(r, 13).HorizontalAlignment = xlLeft
            If r = peerFirst Then ws.Range(ws.cells(r, 1), ws.cells(r, 12)).Font.Color = RR4_ACCENT   ' the target row
            If IsNumeric(Cell(row, 3)) Then ws.cells(r, 4).Font.Color = IIf(CDbl(Cell(row, 3)) >= floorV, CLR_GOOD, CLR_BAD)
            r = r + 1
        End If
    Next row
    Dim peerLast As Long: peerLast = r - 1
    Call WriteNote(ws, r, "EV = market cap + debt + leases + minority - cash & short-term investments.  Peers with negative EBIT or EV/EBIT >= 100 are listed but left out of the band and regression.")
    r = r + 2

    ' ---- thesis + log
    Call Section(ws, r, 1, "500-WORD THESIS  sections [4] numbers . [5] falsification . [6] decision", 10)
    r = r + 1
    For Each row In m_blocks("THESIS")
        ws.cells(r, 1).Value = Cell(row, 0)
        If Left$(Cell(row, 0), 1) = "[" Then ws.cells(r, 1).Font.Color = RR4_ACCENT: ws.cells(r, 1).Font.Bold = True
        r = r + 1
    Next row
    r = r + 1
    Call Section(ws, r, 1, "LOG", 10)
    r = r + 1
    If m_blocks.Exists("LOG") Then
        For Each row In m_blocks("LOG")
            ws.cells(r, 1).Value = Cell(row, 0)
            ws.cells(r, 1).Font.Color = CLR_MUTED
            r = r + 1
        Next row
    End If

    Call DrawScatter(ws, peerFirst, peerLast, leftEnd + 1)
End Sub

Private Sub WriteVals(ws As Worksheet, ByVal r As Long, ByVal c As Long, row As Variant, fmts As Variant)
    Dim i As Long
    For i = 0 To UBound(fmts)
        Dim s As String: s = Cell(row, i)
        With ws.cells(r, c + i)
            If fmts(i) = "@" Then
                .NumberFormat = "@"
                .Value = s
                .HorizontalAlignment = xlLeft
            ElseIf Len(s) = 0 Or Not IsNumeric(s) Then
                .Value = IIf(Len(s) = 0, "-", s)
                .Font.Color = CLR_MUTED
                .HorizontalAlignment = xlRight
            Else
                .Value = CDbl(s)
                .HorizontalAlignment = xlRight
                Select Case fmts(i)
                    Case "pct": .NumberFormat = "0.0%"
                    Case "x": .NumberFormat = "0.0""x"""
                    Case "M": .NumberFormat = "#,##0,,;(#,##0,,);0"
                End Select
            End If
        End With
    Next i
End Sub

Private Sub WriteNote(ws As Worksheet, ByVal r As Long, ByVal txt As String)
    With ws.cells(r, 1)
        .Value = txt
        .Font.Color = CLR_MUTED
        .Font.Italic = True
    End With
End Sub

' EV/EBIT (y) vs ROIC (x): one series per peer row so every point carries
' its own label; the target is orange, peers grey, excluded peers hollow.
' Anchored to the right of the history / peer tables (column N).
Private Sub DrawScatter(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal topRow As Long)
    If r2 < r1 Then Exit Sub
    Dim anchor As Range: Set anchor = ws.cells(topRow, 14)
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(anchor.Left, anchor.Top, 560, 380)
    co.Name = CHART_NAME
    co.Placement = xlMove
    Dim ch As Chart: Set ch = co.Chart
    ch.ChartType = xlXYScatter
    Dim k As Long
    For k = ch.SeriesCollection.count To 1 Step -1: ch.SeriesCollection(k).Delete: Next k
    Dim r As Long, n As Long
    For r = r1 To r2
        If IsNumeric(ws.cells(r, 4).Value) And IsNumeric(ws.cells(r, 5).Value) Then
            If ws.cells(r, 4).Value <> "" And ws.cells(r, 5).Value <> "" Then
                Dim s As Series: Set s = ch.SeriesCollection.NewSeries
                s.Name = CStr(ws.cells(r, 1).Value)
                s.XValues = ws.cells(r, 4)
                s.Values = ws.cells(r, 5)
                s.MarkerStyle = xlMarkerStyleCircle
                s.MarkerSize = IIf(r = r1, 10, 7)
                Dim excluded As Boolean: excluded = (CDbl(ws.cells(r, 5).Value) >= 100)
                If r = r1 Then
                    s.MarkerBackgroundColor = RR4_ACCENT: s.MarkerForegroundColor = RR4_ACCENT
                ElseIf excluded Then
                    s.MarkerBackgroundColor = RGB(0, 0, 0): s.MarkerForegroundColor = RGB(120, 120, 120)
                Else
                    s.MarkerBackgroundColor = RGB(170, 170, 170): s.MarkerForegroundColor = RGB(170, 170, 170)
                End If
                s.HasDataLabels = True
                With s.DataLabels
                    .ShowSeriesName = True: .ShowValue = False: .ShowCategoryName = False
                    .Position = xlLabelPositionRight
                    .Font.Name = FONT_FACE: .Font.Size = 8
                    .Font.Color = IIf(r = r1, RR4_ACCENT, RGB(200, 200, 200))
                End With
                n = n + 1
            End If
        End If
    Next r
    If n = 0 Then co.Delete: Exit Sub
    With ch
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = "EV/EBIT vs ROIC  -  " & Meta("group")
        .ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_TEXT
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
        .ChartTitle.Format.TextFrame2.TextRange.Font.Name = FONT_FACE
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
        With .Axes(xlCategory)
            .HasTitle = True: .AxisTitle.Text = "ROIC (TTM)"
            .AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_MUTED
            .AxisTitle.Format.TextFrame2.TextRange.Font.Size = 8
            .TickLabels.NumberFormat = "0%"
            .TickLabels.Font.Color = CLR_MUTED: .TickLabels.Font.Size = 8
            .Format.Line.ForeColor.RGB = RR4_LINE
            .HasMajorGridlines = True: .MajorGridlines.Format.Line.ForeColor.RGB = RGB(35, 35, 35)
        End With
        With .Axes(xlValue)
            .HasTitle = True: .AxisTitle.Text = "EV / EBIT (TTM)"
            .AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_MUTED
            .AxisTitle.Format.TextFrame2.TextRange.Font.Size = 8
            .TickLabels.NumberFormat = "0""x"""
            .TickLabels.Font.Color = CLR_MUTED: .TickLabels.Font.Size = 8
            .Format.Line.ForeColor.RGB = RR4_LINE
            .HasMajorGridlines = True: .MajorGridlines.Format.Line.ForeColor.RGB = RGB(35, 35, 35)
        End With
    End With
    ' the absolute ROIC floor as a thin vertical guide
    Dim floorV As Double: floorV = Val(Meta("floor"))
    If floorV > 0 Then
        Dim yMax As Double: yMax = ch.Axes(xlValue).MaximumScale
        Dim g As Series: Set g = ch.SeriesCollection.NewSeries
        g.Name = "_floor"
        g.XValues = Array(floorV, floorV)
        g.Values = Array(0, yMax)
        g.ChartType = xlXYScatterLinesNoMarkers
        g.Format.Line.ForeColor.RGB = RGB(90, 90, 90)
        g.Format.Line.DashStyle = msoLineDash
        g.Format.Line.Weight = 1
        g.HasDataLabels = False
    End If
End Sub

' ----------------------------------------------------------------
' Sheet event code (a copy of RR4/SheetValuation_Code.txt), written into
' the document module the first time the page is built.
' ----------------------------------------------------------------
Private Sub EnsureSheetCode(ws As Worksheet)
    On Error GoTo Skip
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 Then
            If comp.Properties("Name").Value = ws.Name Then
                Dim cm As Object: Set cm = comp.CodeModule
                If cm.CountOfLines > 0 Then
                    If InStr(cm.Lines(1, cm.CountOfLines), "ValuationChange") > 0 Then Exit Sub
                    cm.DeleteLines 1, cm.CountOfLines
                End If
                cm.AddFromString "Option Explicit" & vbCrLf & vbCrLf & _
                    "' Valuation page: TICKER / MKT / GROUP input -> modValuationPage.ShowValuation" & vbCrLf & _
                    "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
                    "    Call ValuationChange(Me, Target)" & vbCrLf & _
                    "End Sub" & vbCrLf
                Exit Sub
            End If
        End If
    Next comp
    Exit Sub
Skip:
    Call NavNotify("Valuation page built, but its sheet event code could not be written (enable Trust access to the VBA project object model, or paste RR4/SheetValuation_Code.txt)", True)
End Sub
