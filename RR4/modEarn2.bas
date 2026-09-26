Attribute VB_Name = "modEarn2"
Option Explicit

' ================================================================
'  PEER EARNINGS (nav code E2, E2! = rebuild) - sheet "Peer Earnings"   2026-09-26
' ----------------------------------------------------------------
'  Layer 2 of the Earnings pages ("Industry Read"). Layer 1 is the
'  Earnings (E) page, untouched: the LAYER row here jumps to it with the
'  ticker (click the 1).  This page scores one earnings print on seven
'  items, then lines the ticker up against every member of its segment.
'
'  Data: sheet "EarnData" (hidden - Unhide to bulk-paste), ListObject
'  tblEarn, ONE ROW PER TICKER PER QUARTER (history is never overwritten):
'     TICKER | REPORTED | V1..V7 | HEADLINE | PRINT_n READ_n QUOTE_n SOURCE_n (n = 1..7)
'  The page is a render of that table; everything typed on the page is
'  written straight back into it (a row is created on the first write).
'
'  The 7 items (dropdown vocabulary of each tile):
'     1 Revenue Beat/Inline/Miss . 2 EPS Beat/Inline/Miss
'     3 Rev guide, 4 EPS guide  Raised/Maintained/Lowered/Not guided
'     5 Margin Expansion/Flat/Contraction . 6 Pricing Strong/Neutral/Weak
'     7 Industry Cyclical/Secular/Defensive
'  Symbols: good up-triangle, bad down-triangle, neutral dot, not guided ring.
'  Green = good / red = bad HERE (screenshot template; deliberately unlike
'  the red-up convention of the RR4 pages - same as Library STATUS).
'
'  >>> INDEX and BUCKET are defined in ONE place: VerdictScore / Bucket
'  and the IDX_DIV / BK_* constants just below. The user may send a
'  different formula - change it there and nowhere else. <<<
'
'  Segment (peer group): TW ticker -> first tblGroups (Market = TW) group that
'  holds it, shown through RRG's NormalizeGroupName; US ticker -> first
'  IndustryMap plate (IMAP; the moomoo CSV has plate names only, no parent
'  sector, so it reads "Airlines" not "C. Discretionary - Travel_Airlines").
'  The SEGMENT cell is a dropdown that overrides. Members are matched to
'  tblEarn with the .TW/.TWO suffix stripped; each member's LATEST row is
'  shown (the selected ticker uses the quarter picked in REPORTED).
'
'  Clicks (Worksheet_SelectionChange): a tile or tab = that item goes into
'  the detail strip; a peer's ticker in the table = that ticker becomes the
'  TICKER (same segment); the 1 of LAYER = jump to the E page.
'
'  Geometry: a grid of 14 equal "units" (column B onwards once the nav
'  bar's blank column A is in). Every draw routine addresses cells by
'  (page row, unit) through Cl / Rg, which add the bar's NavOffset / NavLeft,
'  so the same code runs while the page is being built and afterwards.
'  Shapes (peer map) are all named E2_*, placed from cell geometry, xlMove.
'
'  Empty by design until the Bloomberg consensus / guidance data arrives:
'  no sample data is stored. Pure ASCII (VBE import rule).
' ================================================================

Public Const E2_PAGE As String = "Peer Earnings"
Public Const E2_DATA As String = "EarnData"
Public Const E2_TABLE As String = "tblEarn"

' ---------------- INDEX / BUCKET definition (edit here) ----------------
Private Const IDX_DIV As Double = 7             ' INDEX = sum(item scores) / IDX_DIV
Private Const BK_LONG As Double = 0.5           ' INDEX >= this            -> Long
Private Const BK_MONITOR As Double = 0.15       ' INDEX >= this (< LONG)   -> Monitor; below -> Avoid

' --- tblEarn columns ---
Private Const ET_TICKER As Long = 1
Private Const ET_REP As Long = 2
Private Const ET_V1 As Long = 3                 ' V1..V7 = 3..9
Private Const ET_HEAD As Long = 10
Private Const ET_D0 As Long = 10                ' PRINT_n = ET_D0 + 4*(n-1) + 1 ... SOURCE_n = +4
Private Const ET_NCOL As Long = 38
Private Const LIST_COL As Long = 42             ' EarnData: dropdown source lists (segment names, keys, dates)

' --- page rows (page coordinates; + NavOffset once the bar is on) ---
Private Const R_TITLE As Long = 1
Private Const R_TICKER As Long = 2
Private Const R_LAYER As Long = 3
Private Const R_PASTE As Long = 4
Private Const R_SEG As Long = 5
Private Const R_GAP1 As Long = 6
Private Const R_SC As Long = 7
Private Const R_TLBL As Long = 8
Private Const R_TLGAP As Long = 9               ' spacer between the tile labels and the tiles
Private Const R_TILE As Long = 10
Private Const R_GAP2 As Long = 11
Private Const R_DHEAD As Long = 12
Private Const R_PRINT As Long = 13
Private Const R_READ As Long = 14
Private Const R_QUOTE As Long = 15
Private Const R_SRC As Long = 16
Private Const R_GAP3 As Long = 17
Private Const R_MAP As Long = 18
Private Const R_MLBL As Long = 19               ' the key ring block lives on rows 19-24 (left of the axis)
Private Const R_RING1 As Long = 20              ' peer rings sit on rows 20-21
Private Const R_AXIS As Long = 22
Private Const R_TICK As Long = 23
Private Const R_LEG As Long = 24
Private Const R_THDR As Long = 25
Private Const R_THGAP As Long = 26              ' spacer between the table header and the first row
Private Const R_TFIRST As Long = 27

' Logical units: 0 = label column (tickers / block labels), 1..7 = the seven item
' columns (tiles, detail tabs, peer-table verdicts share them), 8 INDEX, 9 BUCKET,
' 10 REPORTED.  Physical columns interleave a narrow black spacer after every
' item column (ColOf), which is what makes the gaps between tiles / tabs / cells.
Private Const NUNIT As Long = 11                ' units 0..10
Private Const NPHYS As Long = 18                ' physical columns behind them
Private Const SP_PTS As Double = 6              ' spacer column width
Private Const SEP_PT As Double = 7.5            ' label -> tile / header -> first row gap (~10 px)
Private Const FONT_FACE As String = "Consolas"
Private Const ITEM_NAME As String = "E2ITEM"    ' hidden workbook name: the item shown in the detail strip

' --- module state (declarations must precede every procedure) ---
Private m_ws As Worksheet
Private m_off As Long
Private m_lc As Long
Private m_im As Variant                         ' IndustryMap A:C values (US segments), Empty = not loaded
Private m_imLoaded As Boolean

Private c_tk As String                          ' upper-cased TICKER cell
Private c_mkt As String                         ' TW / US / CN
Private c_segKey As String
Private c_rep As Double                         ' selected REPORTED serial (0 = none)
Private c_v As Variant                          ' tblEarn body values (Empty when the table is empty)
Private c_n As Long
Private c_selRow As Long                        ' index into c_v of (ticker, quarter), 0 = none
Private c_mem() As String
Private c_nm As Long

Private p_n As Long                             ' peers = segment members
Private p_tk() As String
Private p_row() As Long
Private p_has() As Boolean
Private p_sel() As Boolean
Private p_idx() As Double
Private p_v() As String                         ' (1..p_n, 1..7)
Private p_ord() As Long

' ================================================================
'  Vocabulary, scoring, colours
' ================================================================
Private Function ItemName(ByVal n As Long) As String
    ItemName = Array("Revenue", "EPS", "Rev guide", "EPS guide", "Margin", "Pricing", "Industry")(n - 1)
End Function

Private Function ItemVocab(ByVal n As Long) As String
    ItemVocab = Array("Beat,Inline,Miss", "Beat,Inline,Miss", "Raised,Maintained,Lowered,Not guided", _
                      "Raised,Maintained,Lowered,Not guided", "Expansion,Flat,Contraction", _
                      "Strong,Neutral,Weak", "Cyclical,Secular,Defensive")(n - 1)
End Function

' 0 blank, 1 good, 2 bad, 3 neutral, 4 not guided
Private Function VClass(ByVal v As String) As Long
    Select Case LCase$(Trim$(v))
        Case "", "-": VClass = 0
        Case "beat", "raised", "expansion", "strong": VClass = 1
        Case "miss", "lowered", "contraction", "weak": VClass = 2
        Case "not guided": VClass = 4
        Case Else: VClass = 3
    End Select
End Function

' Score of one verdict (item = 1..7). +1 good, -1 bad, everything else 0
' (blank, neutral, not guided). Per-item weights would go here.
Private Function VerdictScore(ByVal item As Long, ByVal v As String) As Double
    Select Case VClass(v)
        Case 1: VerdictScore = 1
        Case 2: VerdictScore = -1
        Case Else: VerdictScore = 0
    End Select
End Function

Private Function BucketOf(ByVal idx As Double) As String
    If idx >= BK_LONG - 0.000000001 Then
        BucketOf = "Long"
    ElseIf idx >= BK_MONITOR - 0.000000001 Then
        BucketOf = "Monitor"
    Else
        BucketOf = "Avoid"
    End If
End Function

Private Function BucketColor(ByVal b As String) As Long
    Select Case b
        Case "Long": BucketColor = RGB(86, 200, 124)
        Case "Monitor": BucketColor = RGB(235, 150, 40)
        Case Else: BucketColor = RGB(224, 92, 92)
    End Select
End Function

Private Function VSym(ByVal cls As Long) As String
    Select Case cls
        Case 1: VSym = ChrW(&H25B2)
        Case 2: VSym = ChrW(&H25BC)
        Case 3: VSym = ChrW(&H2022)
        Case 4: VSym = ChrW(&H25CB)
    End Select
End Function

' cell number format that shows the symbol in front of the (text) verdict
Private Function VFmt(ByVal cls As Long) As String
    If cls = 0 Then VFmt = "@" Else VFmt = Chr$(34) & VSym(cls) & " " & Chr$(34) & "@"
End Function

Private Function ClsFg(ByVal cls As Long) As Long
    Select Case cls
        Case 1: ClsFg = RGB(86, 200, 124)
        Case 2: ClsFg = RGB(230, 92, 92)
        Case 3: ClsFg = RGB(214, 214, 214)
        Case Else: ClsFg = RGB(140, 140, 140)
    End Select
End Function

Private Function ClsBg(ByVal cls As Long, ByVal tile As Boolean) As Long
    Select Case cls
        Case 1: ClsBg = IIf(tile, RGB(22, 54, 34), RGB(16, 40, 26))
        Case 2: ClsBg = IIf(tile, RGB(56, 24, 24), RGB(42, 18, 18))
        Case Else: ClsBg = IIf(tile, RGB(32, 32, 32), RGB(26, 26, 26))
    End Select
End Function

Private Function ClsBar(ByVal cls As Long) As Long
    Select Case cls
        Case 1: ClsBar = RGB(58, 170, 96)
        Case 2: ClsBar = RGB(190, 64, 64)
        Case Else: ClsBar = RGB(110, 110, 110)
    End Select
End Function

Private Function RingColor(ByVal cls As Long) As Long
    Select Case cls
        Case 1: RingColor = RGB(70, 184, 104)
        Case 2: RingColor = RGB(212, 74, 74)
        Case 3: RingColor = RGB(128, 128, 128)
        Case 4: RingColor = RGB(84, 84, 84)
        Case Else: RingColor = RGB(58, 58, 58)
    End Select
End Function

' ================================================================
'  Small helpers
' ================================================================
Private Function CellStr(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    CellStr = Trim$(CStr(v))
End Function

Private Function NumOr0(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then NumOr0 = CDbl(v): Exit Function       ' IsNumeric(Date) is False
    If IsNumeric(v) Then NumOr0 = CDbl(v)
End Function

' date serial (integer part) of a cell value: Date / number / "yyyy-mm-dd" text; 0 = none
Private Function ParseRep(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then ParseRep = Int(CDbl(v)): Exit Function
    If VarType(v) = vbString Then
        Dim s As String: s = Trim$(CStr(v))
        If s = "" Then Exit Function
        If IsDate(s) Then ParseRep = Int(CDbl(CDate(s)))
        Exit Function
    End If
    If IsNumeric(v) Then
        If CDbl(v) > 30000 Then ParseRep = Int(CDbl(v))
    End If
End Function

Private Function BareKey(ByVal s As String) As String
    s = UCase$(Trim$(s))
    If Right$(s, 4) = ".TWO" Then
        s = Left$(s, Len(s) - 4)
    ElseIf Right$(s, 3) = ".TW" Then
        s = Left$(s, Len(s) - 3)
    End If
    BareKey = Replace(s, "-", ".")
End Function

Private Function MktOf(ByVal tk As String) As String
    Dim u As String: u = UCase$(Trim$(tk))
    If u = "" Then Exit Function
    If Right$(u, 3) = ".TW" Or Right$(u, 4) = ".TWO" Then MktOf = "TW": Exit Function
    If Right$(u, 3) = ".SS" Or Right$(u, 3) = ".SZ" Then MktOf = "CN": Exit Function
    Dim core As String: core = u
    If Len(core) > 1 Then
        If Right$(core, 1) Like "[A-Z]" Then core = Left$(core, Len(core) - 1)
    End If
    If Len(core) >= 4 Then
        If core Like String$(Len(core), "#") Then MktOf = "TW": Exit Function
    End If
    MktOf = "US"
End Function

Private Sub Sync(ByVal ws As Worksheet)
    Set m_ws = ws
    m_off = NavOffset(ws)
    m_lc = NavLeft(ws)
End Sub

' cell at (page row, unit); unit 0 = the first content column
Private Function ColOf(ByVal u As Long) As Long
    If u <= 0 Then
        ColOf = 0
    ElseIf u <= 7 Then
        ColOf = 2 * u - 1
    Else
        ColOf = u + 7
    End If
End Function

' physical column offset (0-based from the first content column) -> unit; -1 = spacer / outside
Private Function UnitOfCol(ByVal k As Long) As Long
    UnitOfCol = -1
    If k = 0 Then
        UnitOfCol = 0
    ElseIf k >= 1 And k <= 13 Then
        If k Mod 2 = 1 Then UnitOfCol = (k + 1) \ 2
    ElseIf k >= 15 And k <= 17 Then
        UnitOfCol = k - 7
    End If
End Function

Private Function Cl(ByVal pr As Long, ByVal u As Long) As Range
    Set Cl = m_ws.Cells(pr + m_off, m_lc + 1 + ColOf(u))
End Function

Private Function Rg(ByVal pr As Long, ByVal u1 As Long, ByVal u2 As Long) As Range
    Set Rg = m_ws.Range(Cl(pr, u1), Cl(pr, u2))
End Function

Private Sub Edge(ByVal rng As Range, ByVal which As Long, ByVal clr As Long, ByVal wt As Long)
    With rng.Borders(which)
        .LineStyle = xlContinuous
        .Color = clr
        .Weight = wt
    End With
End Sub

Private Sub NoEdge(ByVal rng As Range, ByVal which As Long)
    rng.Borders(which).LineStyle = xlNone
End Sub

Private Function CurItem() As Long
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(ITEM_NAME).RefersTo
    On Error GoTo 0
    s = Replace(s, "=", "")
    CurItem = CLng(Val(s))
    If CurItem < 1 Or CurItem > 7 Then CurItem = 1
End Function

Private Sub SetItem(ByVal n As Long)
    On Error Resume Next
    ThisWorkbook.Names(ITEM_NAME).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=ITEM_NAME, RefersTo:="=" & n, Visible:=False
End Sub

Private Function DetCol(ByVal item As Long, ByVal k As Long) As Long
    DetCol = ET_D0 + 4 * (item - 1) + k              ' k: 1 print, 2 read, 3 quote, 4 source
End Function

' ================================================================
'  Data sheet + table
' ================================================================
Private Function EarnTable() As ListObject
    On Error Resume Next
    Set EarnTable = ThisWorkbook.Worksheets(E2_DATA).ListObjects(E2_TABLE)
End Function

Private Function TableRows(ByVal lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    TableRows = lo.DataBodyRange.Rows.Count
End Function

Private Function EnsureData() As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(E2_DATA)
    On Error GoTo 0
    Dim fresh As Boolean
    If ws Is Nothing Then
        Dim prev As Object: Set prev = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = E2_DATA
        ActiveWindow.DisplayGridlines = False
        fresh = True
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If

    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = FONT_FACE: .Font.Size = 9: .Font.Color = RGB(221, 221, 221)
        .VerticalAlignment = xlCenter
    End With
    With ws.Cells(1, 1)
        .Value = "EARN DATA  .  data for the Peer Earnings page (E2) - one row per TICKER per REPORTED quarter; V1..V7 = Revenue, EPS, Rev guide, EPS guide, Margin, Pricing, Industry verdicts; the page writes here as you type"
        .Font.Color = RR4_ACCENT: .Font.Bold = True
    End With

    Dim lo As ListObject: Set lo = EarnTable()
    Dim hdr(0 To ET_NCOL - 1) As String
    Dim j As Long, n As Long
    hdr(0) = "TICKER": hdr(1) = "REPORTED"
    For n = 1 To 7: hdr(1 + n) = "V" & n: Next n
    hdr(9) = "HEADLINE"
    For n = 1 To 7
        hdr(ET_D0 + 4 * (n - 1) + 0) = "PRINT_" & n
        hdr(ET_D0 + 4 * (n - 1) + 1) = "READ_" & n
        hdr(ET_D0 + 4 * (n - 1) + 2) = "QUOTE_" & n
        hdr(ET_D0 + 4 * (n - 1) + 3) = "SOURCE_" & n
    Next n
    If lo Is Nothing Then
        For j = 0 To ET_NCOL - 1: ws.Cells(3, j + 1).Value = hdr(j): Next j
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(3, 1), ws.Cells(4, ET_NCOL)), , xlYes)
        lo.Name = E2_TABLE
        ws.Range(ws.Cells(4, 1), ws.Cells(4, ET_NCOL)).NumberFormat = "@"
        ws.Cells(4, ET_REP).NumberFormat = "yyyy-mm-dd"
    End If
    lo.TableStyle = ""
    With lo.HeaderRowRange
        .Font.Color = RR4_ACCENT: .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RR4_LINE
    End With
    ws.Columns(ET_TICKER).ColumnWidth = 12
    ws.Columns(ET_REP).ColumnWidth = 12
    For j = ET_V1 To ET_V1 + 6: ws.Columns(j).ColumnWidth = 12: Next j
    ws.Columns(ET_HEAD).ColumnWidth = 44
    For n = 1 To 7
        ws.Columns(DetCol(n, 1)).ColumnWidth = 30
        ws.Columns(DetCol(n, 2)).ColumnWidth = 44
        ws.Columns(DetCol(n, 3)).ColumnWidth = 44
        ws.Columns(DetCol(n, 4)).ColumnWidth = 30
    Next n
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.Interior.Color = RGB(8, 8, 8)
        lo.DataBodyRange.Font.Color = RGB(221, 221, 221)
    End If
    ws.Cells(3, LIST_COL).Value = "SEG NAME (auto)"
    ws.Cells(3, LIST_COL + 1).Value = "SEG KEY (auto)"
    ws.Cells(3, LIST_COL + 2).Value = "DATES (auto)"
    ws.Range(ws.Cells(3, LIST_COL), ws.Cells(3, LIST_COL + 2)).Font.Color = RR4_LINE
    ws.Columns(LIST_COL).ColumnWidth = 30
    ws.Columns(LIST_COL + 1).ColumnWidth = 14
    ws.Columns(LIST_COL + 2).ColumnWidth = 12

    Call WriteSheetCode(ws, "Option Explicit" & vbCrLf & vbCrLf & _
        "' EarnData (modEarn2): editing tblEarn redraws the Peer Earnings page" & vbCrLf & _
        "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    Call Earn2DataChange(Me, Target)" & vbCrLf & _
        "End Sub" & vbCrLf, "Earn2DataChange")
    If fresh Then ws.Visible = xlSheetHidden
    Set EnsureData = lo
End Function

Private Sub LoadTable()
    c_n = 0
    c_v = Empty
    Dim lo As ListObject: Set lo = EarnTable()
    If TableRows(lo) = 0 Then Exit Sub
    c_v = lo.DataBodyRange.Value
    c_n = UBound(c_v, 1)
End Sub

Private Function FindRow(ByVal bare As String, ByVal rep As Double) As Long
    Dim i As Long
    For i = 1 To c_n
        If BareKey(CellStr(c_v(i, ET_TICKER))) = bare Then
            If Int(NumOr0(c_v(i, ET_REP))) = rep Then FindRow = i: Exit Function
        End If
    Next i
End Function

Private Function LatestRow(ByVal bare As String) As Long
    Dim i As Long, best As Double, d As Double
    best = -1
    For i = 1 To c_n
        If BareKey(CellStr(c_v(i, ET_TICKER))) = bare Then
            d = NumOr0(c_v(i, ET_REP))
            If d >= best Then best = d: LatestRow = i
        End If
    Next i
End Function

' Row (1-based, table body) for (ticker, quarter): found or created.
Private Function EnsureRowFor(ByVal tk As String, ByVal rep As Double) As Long
    Dim lo As ListObject: Set lo = EarnTable()
    If lo Is Nothing Then Set lo = EnsureData()
    Call LoadTable
    Dim r As Long: r = FindRow(BareKey(tk), rep)
    If r > 0 Then EnsureRowFor = r: Exit Function
    If TableRows(lo) = 0 Then lo.Resize lo.Range.Resize(2)
    Dim i As Long, blank As Long
    For i = 1 To lo.DataBodyRange.Rows.Count
        If CellStr(lo.DataBodyRange.Cells(i, ET_TICKER).Value) = "" Then blank = i: Exit For
    Next i
    If blank = 0 Then
        lo.Resize lo.Range.Resize(lo.Range.Rows.Count + 1)
        blank = lo.DataBodyRange.Rows.Count
    End If
    With lo.DataBodyRange
        .Range(.Cells(blank, 1), .Cells(blank, ET_NCOL)).NumberFormat = "@"
        .Cells(blank, ET_REP).NumberFormat = "yyyy-mm-dd"
        .Cells(blank, ET_TICKER).Value = UCase$(Trim$(tk))
        .Cells(blank, ET_REP).Value2 = rep
        .Rows(blank).Interior.Color = RGB(8, 8, 8)
        .Rows(blank).Font.Color = RGB(221, 221, 221)
    End With
    Call LoadTable
    EnsureRowFor = blank
End Function

Private Sub SetField(ByVal row As Long, ByVal col As Long, ByVal txt As String)
    Dim lo As ListObject: Set lo = EarnTable()
    If lo Is Nothing Then Exit Sub
    With lo.DataBodyRange.Cells(row, col)
        .NumberFormat = "@"
        .Value = txt
    End With
End Sub

' ================================================================
'  Segments
' ================================================================
Private Sub LoadIM()
    m_imLoaded = True
    m_im = Empty
    Dim wm As Worksheet
    On Error Resume Next
    Set wm = ThisWorkbook.Worksheets("IndustryMap")
    On Error GoTo 0
    If wm Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = wm.Cells(wm.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub
    m_im = wm.Range(wm.Cells(2, 1), wm.Cells(lastR, 3)).Value
End Sub

' Segment list of a market: display names + keys (TW: raw tblGroups name;
' US: IndustryMap plateCode). n = 0 when the source is missing / empty.
Private Sub SegList(ByVal mkt As String, ByRef disp() As String, ByRef key() As String, ByRef n As Long)
    n = 0
    Dim i As Long
    If mkt = "TW" Then
        Dim gn As Variant: gn = GetGroupNames("TW")
        If IsEmpty(gn) Then Exit Sub
        Dim used As Object: Set used = CreateObject("Scripting.Dictionary")
        used.CompareMode = vbTextCompare
        ReDim disp(1 To UBound(gn) + 1): ReDim key(1 To UBound(gn) + 1)
        Dim nm As String
        For i = 0 To UBound(gn)
            nm = NormalizeGroupName(CStr(gn(i)))
            If used.Exists(nm) Then
                used(nm) = used(nm) + 1
                nm = nm & " (" & used(nm) & ")"
            Else
                used(nm) = 1
            End If
            n = n + 1
            disp(n) = nm: key(n) = CStr(gn(i))
        Next i
    ElseIf mkt = "US" Then
        If Not m_imLoaded Then Call LoadIM
        If IsEmpty(m_im) Then Exit Sub
        Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
        Dim names As Object: Set names = CreateObject("Scripting.Dictionary")
        names.CompareMode = vbTextCompare
        ReDim disp(1 To UBound(m_im, 1)): ReDim key(1 To UBound(m_im, 1))
        Dim code As String, pn As String
        For i = 1 To UBound(m_im, 1)
            code = CellStr(m_im(i, 1))
            If code <> "" Then
                If Not seen.Exists(code) Then
                    seen(code) = True
                    pn = CellStr(m_im(i, 2))
                    If pn = "" Then pn = code
                    If names.Exists(pn) Then pn = pn & " (" & code & ")" Else names(pn) = True
                    n = n + 1
                    disp(n) = pn: key(n) = code
                End If
            End If
        Next i
    End If
End Sub

Private Function SegDispOf(ByVal mkt As String, ByVal key As String) As String
    Dim disp() As String, keys() As String, n As Long, i As Long
    Call SegList(mkt, disp, keys, n)
    For i = 1 To n
        If keys(i) = key Then SegDispOf = disp(i): Exit Function
    Next i
End Function

' typed SEGMENT text -> key ("" = none): exact display / exact key, then a unique partial hit
Private Function SegKeyOf(ByVal mkt As String, ByVal txt As String) As String
    txt = Trim$(txt)
    If txt = "" Then Exit Function
    Dim disp() As String, keys() As String, n As Long, i As Long, hits As Long, hitKey As String
    Call SegList(mkt, disp, keys, n)
    For i = 1 To n
        If StrComp(disp(i), txt, vbTextCompare) = 0 Or StrComp(keys(i), txt, vbTextCompare) = 0 Then
            SegKeyOf = keys(i): Exit Function
        End If
    Next i
    For i = 1 To n
        If InStr(1, disp(i), txt, vbTextCompare) > 0 Then hits = hits + 1: hitKey = keys(i)
    Next i
    If hits = 1 Then SegKeyOf = hitKey
End Function

' members of a segment as typed forms, deduped by bare code
Private Sub SegMembers(ByVal mkt As String, ByVal key As String)
    c_nm = 0
    ReDim c_mem(1 To 1)
    If key = "" Then Exit Sub
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, s As String, b As String
    If mkt = "TW" Then
        Dim tks As Variant: tks = GetSectorTickers("TW", key)
        If IsEmpty(tks) Then Exit Sub
        ReDim c_mem(1 To UBound(tks) - LBound(tks) + 1)
        For i = LBound(tks) To UBound(tks)
            s = UCase$(Trim$(CStr(tks(i))))
            b = BareKey(s)
            If s <> "" And Not seen.Exists(b) Then seen(b) = True: c_nm = c_nm + 1: c_mem(c_nm) = s
        Next i
    ElseIf mkt = "US" Then
        If Not m_imLoaded Then Call LoadIM
        If IsEmpty(m_im) Then Exit Sub
        ReDim c_mem(1 To UBound(m_im, 1))
        For i = 1 To UBound(m_im, 1)
            If CellStr(m_im(i, 1)) = key Then
                s = UCase$(CellStr(m_im(i, 3)))
                b = BareKey(s)
                If s <> "" And Not seen.Exists(b) Then seen(b) = True: c_nm = c_nm + 1: c_mem(c_nm) = s
            End If
        Next i
    End If
End Sub

' first segment (list order) whose members hold the ticker; "" = none
Private Function AutoSegKey(ByVal mkt As String, ByVal tk As String) As String
    Dim bare As String: bare = BareKey(tk)
    Dim i As Long
    If mkt = "TW" Then
        Dim gn As Variant, tks As Variant, j As Long
        gn = GetGroupNames("TW")
        If IsEmpty(gn) Then Exit Function
        For i = 0 To UBound(gn)
            tks = GetSectorTickers("TW", CStr(gn(i)))
            If Not IsEmpty(tks) Then
                For j = LBound(tks) To UBound(tks)
                    If BareKey(CStr(tks(j))) = bare Then AutoSegKey = CStr(gn(i)): Exit Function
                Next j
            End If
        Next i
    ElseIf mkt = "US" Then
        If Not m_imLoaded Then Call LoadIM
        If IsEmpty(m_im) Then Exit Function
        For i = 1 To UBound(m_im, 1)
            If BareKey(CellStr(m_im(i, 3))) = bare Then AutoSegKey = CellStr(m_im(i, 1)): Exit Function
        Next i
    End If
End Function

' does segment key hold the ticker?
Private Function SegHolds(ByVal mkt As String, ByVal key As String, ByVal tk As String) As Boolean
    If key = "" Then Exit Function
    Call SegMembers(mkt, key)
    Dim i As Long, bare As String: bare = BareKey(tk)
    For i = 1 To c_nm
        If BareKey(c_mem(i)) = bare Then SegHolds = True: Exit Function
    Next i
End Function

' Rewrites the SEGMENT dropdown source (E2SegList) for a market.
Private Sub WriteSegList(ByVal mkt As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(E2_DATA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.Range(ws.Cells(4, LIST_COL), ws.Cells(4 + 2000, LIST_COL + 1)).ClearContents
    Dim disp() As String, keys() As String, n As Long, i As Long
    Call SegList(mkt, disp, keys, n)
    If n > 0 Then
        Dim arr() As Variant: ReDim arr(1 To n, 1 To 2)
        For i = 1 To n: arr(i, 1) = disp(i): arr(i, 2) = keys(i): Next i
        ws.Range(ws.Cells(4, LIST_COL), ws.Cells(3 + n, LIST_COL + 1)).NumberFormat = "@"
        ws.Range(ws.Cells(4, LIST_COL), ws.Cells(3 + n, LIST_COL + 1)).Value = arr
    End If
    Dim lastR As Long: lastR = IIf(n > 0, 3 + n, 4)
    On Error Resume Next
    ThisWorkbook.Names("E2SegList").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="E2SegList", RefersTo:="='" & E2_DATA & "'!" & _
        ws.Range(ws.Cells(4, LIST_COL), ws.Cells(lastR, LIST_COL)).Address, Visible:=False
End Sub

' Rewrites the REPORTED dropdown source (E2DateList): this ticker's quarters, newest first.
Private Sub WriteDateList(ByVal tk As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(E2_DATA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.Range(ws.Cells(4, LIST_COL + 2), ws.Cells(4 + 200, LIST_COL + 2)).ClearContents
    Dim ds() As Double, n As Long, i As Long, j As Long, t As Double
    ReDim ds(1 To 201)
    Dim bare As String: bare = BareKey(tk)
    If bare <> "" Then
        Call LoadTable
        For i = 1 To c_n
            If BareKey(CellStr(c_v(i, ET_TICKER))) = bare Then
                t = Int(NumOr0(c_v(i, ET_REP)))
                If t > 0 And n < 200 Then
                    Dim dup As Boolean: dup = False
                    For j = 1 To n
                        If ds(j) = t Then dup = True: Exit For
                    Next j
                    If Not dup Then n = n + 1: ds(n) = t
                End If
            End If
        Next i
    End If
    For i = 1 To n - 1                                   ' newest first
        For j = i + 1 To n
            If ds(j) > ds(i) Then t = ds(i): ds(i) = ds(j): ds(j) = t
        Next j
    Next i
    For i = 1 To n
        ws.Cells(3 + i, LIST_COL + 2).NumberFormat = "yyyy-mm-dd"
        ws.Cells(3 + i, LIST_COL + 2).Value2 = ds(i)
    Next i
    Dim lastR As Long: lastR = IIf(n > 0, 3 + n, 4)
    On Error Resume Next
    ThisWorkbook.Names("E2DateList").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="E2DateList", RefersTo:="='" & E2_DATA & "'!" & _
        ws.Range(ws.Cells(4, LIST_COL + 2), ws.Cells(lastR, LIST_COL + 2)).Address, Visible:=False
End Sub

' ================================================================
'  Page sheet + shell
' ================================================================
Private Function EnsurePage() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(E2_PAGE)
    On Error GoTo 0
    If ws Is Nothing Then
        Dim prev As Object: Set prev = ActiveSheet
        Dim after As Object
        On Error Resume Next
        Set after = ThisWorkbook.Worksheets(NavSheetName("E"))
        On Error GoTo 0
        If after Is Nothing Then Set after = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set ws = ThisWorkbook.Worksheets.Add(After:=after)
        ws.Name = E2_PAGE
        ActiveWindow.DisplayGridlines = False
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If
    Set EnsurePage = ws
End Function

Private Sub SetColPts(ByVal col As Range, ByVal pts As Double)
    col.ColumnWidth = pts / 5.4
    Dim w As Double: w = col.Width
    If w > 0 Then col.ColumnWidth = col.ColumnWidth * pts / w
End Sub

Private Sub DeleteE2Shapes(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "E2_" Then ws.Shapes(i).Delete
    Next i
End Sub

' E2! - rebuild the whole page (inputs are kept; nothing is fetched except the company name).
Public Sub BuildEarn2()
    Dim lo As ListObject: Set lo = EnsureData()
    Dim ws As Worksheet: Set ws = EnsurePage()
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fail

    ' read what the page holds now (bar geometry as it is)
    Dim tk As String, segTxt As String, repV As Variant, pasteTxt As String, nm As String
    If NavHasRows(ws) Then
        Call Sync(ws)
        tk = UCase$(CellStr(Cl(R_TICKER, 1).Value))
        segTxt = CellStr(Cl(R_SEG, 1).Value)
        repV = Cl(R_SEG, 5).Value
        pasteTxt = CellStr(Cl(R_PASTE, 1).Value)
        nm = CellStr(Cl(R_TICKER, 3).Value)
    End If

    Call NavStrip(ws)
    Call DeleteE2Shapes(ws)
    ws.UsedRange.UnMerge
    ws.Cells.Clear
    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = FONT_FACE
        .Font.Size = 9
        .Font.Color = RGB(221, 221, 221)
        .VerticalAlignment = xlCenter
        .RowHeight = 17
    End With
    Call NavAdd(ws, "E2")
    Call Sync(ws)

    Call DrawShell(ws)
    Cl(R_TICKER, 1).NumberFormat = "@"
    Cl(R_TICKER, 1).Value = tk
    Cl(R_TICKER, 3).Value = nm
    Cl(R_SEG, 1).Value = segTxt
    If ParseRep(repV) > 0 Then Cl(R_SEG, 5).Value2 = ParseRep(repV)
    Cl(R_PASTE, 1).Value = pasteTxt
    m_imLoaded = False
    Call WriteSegList(MktOf(tk))
    Call WriteDateList(tk)
    Call AddValidations
    If tk <> "" And segTxt = "" Then Call AutoFillSegment(tk)
    Call LoadCtx(ws)
    Call DrawContent(True)

    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
    Call NavGoto("E2", ws)
    Call FinishSheetCode(ws)
    Call NavNotify("PEER EARNINGS built" & IIf(TableRows(lo) = 0, " - tblEarn is empty (fill it from the page or paste into EarnData)", ""))
    Exit Sub
Fail:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    Call NavNotify("PEER EARNINGS error: " & msg, True)
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
End Sub

Private Sub FinishSheetCode(ByVal ws As Worksheet)
    Call WriteSheetCode(ws, "Option Explicit" & vbCrLf & vbCrLf & _
        "' Peer Earnings page (modEarn2): typed inputs + tile / tab / peer clicks" & vbCrLf & _
        "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    Call Earn2Change(Me, Target)" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & _
        "    Call Earn2Select(Me, Target)" & vbCrLf & _
        "End Sub" & vbCrLf, "Earn2Select")
End Sub

Private Sub StackLabel(ByVal c As Range, ByVal txt As String)
    With c
        .Value = txt
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 8
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub DimLabel(ByVal c As Range, ByVal txt As String)
    With c
        .Value = txt
        .Font.Color = RGB(130, 130, 130)
        .Font.Size = 8
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub DrawShell(ByVal ws As Worksheet)
    Dim u As Long, i As Long
    Dim wpts As Variant
    wpts = Array(72, 100, SP_PTS, 100, SP_PTS, 100, SP_PTS, 100, SP_PTS, 100, SP_PTS, 100, SP_PTS, 100, SP_PTS, 60, 70, 84)
    For u = 0 To NPHYS - 1
        Call SetColPts(ws.Columns(m_lc + 1 + u), CDbl(wpts(u)))
    Next u
    ws.Rows(R_GAP1 + m_off).RowHeight = 8
    ws.Rows(R_TITLE + m_off).RowHeight = 22
    ws.Rows(R_TICKER + m_off).RowHeight = 22
    ws.Rows(R_TLGAP + m_off).RowHeight = SEP_PT
    ws.Rows(R_TILE + m_off).RowHeight = 32
    ws.Rows(R_THGAP + m_off).RowHeight = SEP_PT
    ws.Rows(R_GAP2 + m_off).RowHeight = 10
    ws.Rows(R_READ + m_off).RowHeight = 34
    ws.Rows(R_QUOTE + m_off).RowHeight = 34
    ws.Rows(R_GAP3 + m_off).RowHeight = 12

    ' --- title + ticker row ---
    With Cl(R_TITLE, 0)
        .Value = "EARNINGS"
        .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 12
    End With
    Call StackLabel(Cl(R_TICKER, 0), "TICKER")
    With Cl(R_TICKER, 1)
        .NumberFormat = "@"
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True: .Font.Size = 11
        .HorizontalAlignment = xlLeft
        Call Edge(Cl(R_TICKER, 1), xlEdgeLeft, RGB(235, 235, 235), xlThin)
        Call Edge(Cl(R_TICKER, 1), xlEdgeRight, RGB(235, 235, 235), xlThin)
        Call Edge(Cl(R_TICKER, 1), xlEdgeTop, RGB(235, 235, 235), xlThin)
        Call Edge(Cl(R_TICKER, 1), xlEdgeBottom, RGB(235, 235, 235), xlThin)
    End With
    With Cl(R_TICKER, 3)
        .Font.Bold = True: .Font.Size = 12
        .Font.Color = RGB(245, 245, 245)
        .HorizontalAlignment = xlLeft
    End With

    ' --- LAYER / PASTE / SEGMENT ---
    Call StackLabel(Cl(R_LAYER, 0), "LAYER")
    With Cl(R_LAYER, 1)
        .Value = "1"
        .Font.Color = RGB(130, 130, 130): .Font.Size = 10
        .HorizontalAlignment = xlRight
    End With
    With Cl(R_LAYER, 2)
        .Value = "2  Industry Read"
        .Font.Color = RGB(214, 214, 214): .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .Characters(1, 1).Font.Color = RR4_ACCENT
        .Characters(1, 1).Font.Bold = True
        .Characters(1, 1).Font.Size = 11
    End With
    Call Edge(Cl(R_LAYER, 2), xlEdgeLeft, RR4_LINE, xlThin)
    Call StackLabel(Cl(R_PASTE, 0), "PASTE")
    With Rg(R_PASTE, 1, 2)
        .Merge
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .HorizontalAlignment = xlLeft
    End With
    Call StackLabel(Cl(R_SEG, 0), "SEGMENT")
    With Rg(R_SEG, 1, 3)
        .Merge
        .NumberFormat = "@"
        .HorizontalAlignment = xlLeft
        .Font.Name = "Noto Sans TC"
        .Font.Color = RGB(235, 235, 235)
    End With
    Call StackLabel(Cl(R_SEG, 4), "REPORTED")
    With Cl(R_SEG, 5)
        .NumberFormat = "yyyy-mm-dd"
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' --- summary block labels (values are dynamic) ---
    Dim bl As Variant: bl = Array("MARGINS", "GUIDES", "BUCKETS", "AVG INDEX", "REPORTED")
    Dim bu As Variant: bu = Array(3, 5, 7, 9, 10)                ' first unit of each block
    For i = 0 To 4
        Call StackLabel(Cl(R_LAYER, CLng(bu(i))), CStr(bl(i)))
        Cl(R_LAYER, CLng(bu(i))).Font.Size = 7
        Call Edge(Cl(R_LAYER, CLng(bu(i))), xlEdgeLeft, RR4_LINE, xlThin)
        Call Edge(Cl(R_PASTE, CLng(bu(i))), xlEdgeLeft, RR4_LINE, xlThin)
        Cl(R_LAYER, CLng(bu(i))).IndentLevel = 1
        Cl(R_PASTE, CLng(bu(i))).IndentLevel = 1
    Next i

    ' --- SCORECARD ---
    Call Edge(Rg(R_SC, 0, NUNIT - 1), xlEdgeTop, RR4_LINE, xlThin)
    Call StackLabel(Cl(R_SC, 0), "SCORECARD")
    With Rg(R_SC, 1, NUNIT - 1)
        .Merge
        .NumberFormat = "@"
        .HorizontalAlignment = xlLeft
        .Font.Color = RGB(200, 200, 200)
        .IndentLevel = 1
    End With
    For i = 1 To 7
        With Cl(R_TLBL, i)
            .Value = i & " " & ItemName(i)
            .Font.Color = RR4_ACCENT: .Font.Bold = True: .Font.Size = 7
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
        End With
        With Cl(R_TILE, i)
            .NumberFormat = "@"
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            .Font.Size = 10
        End With
    Next i

    ' --- detail strip ---
    Call Edge(Rg(R_DHEAD, 0, NUNIT - 1), xlEdgeTop, RR4_LINE, xlThin)
    Cl(R_DHEAD, 0).Font.Color = RR4_ACCENT
    Cl(R_DHEAD, 0).Font.Bold = True
    Call DimLabel(Cl(R_PRINT, 0), "Print vs bar")
    Call DimLabel(Cl(R_READ, 0), "Read")
    Call DimLabel(Cl(R_QUOTE, 0), "Quote")
    Call DimLabel(Cl(R_SRC, 0), "Source")
    Dim rr As Variant
    For Each rr In Array(R_PRINT, R_READ, R_QUOTE, R_SRC)
        With Rg(CLng(rr), 1, NUNIT - 1)
            .Merge
            .NumberFormat = "@"
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .Font.Color = RGB(214, 214, 214)
            .Font.Size = 9
        End With
        Cl(CLng(rr), 0).VerticalAlignment = xlTop
    Next rr

    ' --- peer map + table static parts ---
    Call Edge(Rg(R_MAP, 0, NUNIT - 1), xlEdgeTop, RR4_LINE, xlThin)
    Call StackLabel(Cl(R_MAP, 0), "PEER MAP")
    Cl(R_MAP, 0).Font.Size = 9
    For i = 1 To 7
        With Cl(R_THDR, i)
            .Value = i & " " & ItemName(i)
            .Font.Color = RR4_ACCENT: .Font.Size = 7: .Font.Bold = True
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        End With
    Next i
    With Cl(R_THDR, 8)
        .Value = "INDEX": .Font.Color = RR4_ACCENT: .Font.Size = 7: .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With Cl(R_THDR, 9)
        .Value = "BUCKET": .Font.Color = RR4_ACCENT: .Font.Size = 7: .Font.Bold = True
        .HorizontalAlignment = xlLeft: .IndentLevel = 1
    End With
    With Cl(R_THDR, 10)
        .Value = "REPORTED": .Font.Color = RR4_ACCENT: .Font.Size = 7: .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

' dropdowns: SEGMENT (name list of the ticker's market), REPORTED (this ticker's quarters), 7 tiles
Private Sub AddValidations()
    Dim i As Long
    On Error Resume Next
    With Cl(R_SEG, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="=E2SegList"
        .IgnoreBlank = True
        .ShowError = False
    End With
    With Cl(R_SEG, 5).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="=E2DateList"
        .IgnoreBlank = True
        .ShowError = False
    End With
    For i = 1 To 7
        With Cl(R_TILE, i).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=ItemVocab(i)
            .IgnoreBlank = True
            .ShowError = False
        End With
    Next i
    On Error GoTo 0
End Sub

' ================================================================
'  Context (read the page inputs + tblEarn, resolve the peers)
' ================================================================
Private Sub LoadCtx(ByVal ws As Worksheet)
    Call Sync(ws)
    Call LoadTable
    c_tk = UCase$(CellStr(Cl(R_TICKER, 1).Value))
    c_mkt = MktOf(c_tk)
    c_rep = ParseRep(Cl(R_SEG, 5).Value)
    c_selRow = 0
    c_segKey = ""
    c_nm = 0
    ReDim c_mem(1 To 1)
    If c_tk <> "" Then
        m_imLoaded = False
        c_segKey = SegKeyOf(c_mkt, CellStr(Cl(R_SEG, 1).Value))
        Call SegMembers(c_mkt, c_segKey)
        If c_nm = 0 Then
            c_nm = 1
            ReDim c_mem(1 To 1)
            c_mem(1) = c_tk
        End If
        If c_rep > 0 Then
            c_selRow = FindRow(BareKey(c_tk), c_rep)
        Else
            c_selRow = LatestRow(BareKey(c_tk))
            If c_selRow > 0 Then c_rep = Int(NumOr0(c_v(c_selRow, ET_REP)))
        End If
    End If
    Call BuildPeers
End Sub

Private Sub BuildPeers()
    p_n = c_nm
    If p_n < 1 Then p_n = 0
    If p_n = 0 Then Exit Sub
    ReDim p_tk(1 To p_n): ReDim p_row(1 To p_n): ReDim p_has(1 To p_n)
    ReDim p_sel(1 To p_n): ReDim p_idx(1 To p_n): ReDim p_v(1 To p_n, 1 To 7): ReDim p_ord(1 To p_n)
    Dim i As Long, k As Long, r As Long, bare As String, anyV As Boolean, sc As Double
    Dim selBare As String: selBare = BareKey(c_tk)
    For i = 1 To p_n
        p_tk(i) = c_mem(i)
        bare = BareKey(c_mem(i))
        p_sel(i) = (bare = selBare)
        If p_sel(i) And c_selRow > 0 Then r = c_selRow Else r = LatestRow(bare)
        p_row(i) = r
        anyV = False: sc = 0
        If r > 0 Then
            For k = 1 To 7
                p_v(i, k) = CellStr(c_v(r, ET_V1 + k - 1))
                If p_v(i, k) = "-" Then p_v(i, k) = ""
                If p_v(i, k) <> "" Then anyV = True
                sc = sc + VerdictScore(k, p_v(i, k))
            Next k
        End If
        p_has(i) = anyV
        If anyV Then p_idx(i) = sc / IDX_DIV
    Next i
    ' order: peers with data by INDEX descending (stable), the rest in member order
    Dim n As Long, j As Long, t As Long
    For i = 1 To p_n
        If p_has(i) Then n = n + 1: p_ord(n) = i
    Next i
    For i = 2 To n                                       ' insertion sort: stable, INDEX descending
        t = p_ord(i): j = i - 1
        Do While j >= 1
            If p_idx(p_ord(j)) < p_idx(t) - 0.0000000001 Then
                p_ord(j + 1) = p_ord(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        p_ord(j + 1) = t
    Next i
    For i = 1 To p_n
        If Not p_has(i) Then n = n + 1: p_ord(n) = i
    Next i
End Sub

' ================================================================
'  Content (everything dynamic)
' ================================================================
Private Sub DrawContent(ByVal withCompany As Boolean)
    Call DrawSummary
    Call DrawTiles
    Call DrawDetail
    Call DrawMapAndTable
End Sub

Private Sub PaintBlack(ByVal rng As Range)
    rng.Interior.Color = RGB(0, 0, 0)
    rng.Font.Color = RGB(221, 221, 221)
    rng.Font.Bold = False
End Sub

Private Sub DrawSummary()
    Dim i As Long, up As Long, dn As Long, rs As Long, cut As Long
    Dim nL As Long, nM As Long, nA As Long, nData As Long, sumIdx As Double
    Dim dMin As Double, dMax As Double
    dMin = 0: dMax = 0
    For i = 1 To p_n
        If p_has(i) Then
            nData = nData + 1
            sumIdx = sumIdx + p_idx(i)
            Select Case LCase$(p_v(i, 5))
                Case "expansion": up = up + 1
                Case "contraction": dn = dn + 1
            End Select
            Dim k As Long
            For k = 3 To 4
                Select Case LCase$(p_v(i, k))
                    Case "raised": rs = rs + 1
                    Case "lowered": cut = cut + 1
                End Select
            Next k
            Select Case BucketOf(p_idx(i))
                Case "Long": nL = nL + 1
                Case "Monitor": nM = nM + 1
                Case Else: nA = nA + 1
            End Select
            Dim d As Double: d = Int(NumOr0(c_v(p_row(i), ET_REP)))
            If d > 0 Then
                If dMin = 0 Or d < dMin Then dMin = d
                If d > dMax Then dMax = d
            End If
        End If
    Next i
    Dim sepTxt As String: sepTxt = " " & ChrW(&HB7) & " "
    Dim green As Long, red As Long, orange As Long, dimc As Long
    green = RGB(86, 200, 124): red = RGB(224, 92, 92): orange = RGB(235, 150, 40): dimc = RGB(120, 120, 120)

    ' MARGINS
    Call TwoPart(Cl(R_PASTE, 3), up & " up", green, dn & " down", red, sepTxt, dimc)
    ' GUIDES
    Call TwoPart(Cl(R_PASTE, 5), rs & " raised", green, cut & " cut", red, sepTxt, dimc)
    ' BUCKETS: long first, then monitor, then avoid - non-zero parts only
    With Cl(R_PASTE, 7)
        .Font.Size = 10: .Font.Bold = False: .HorizontalAlignment = xlLeft
        Dim txt As String, parts(1 To 3) As String, cols(1 To 3) As Long, np As Long
        If nL > 0 Then np = np + 1: parts(np) = nL & " long": cols(np) = green
        If nM > 0 Then np = np + 1: parts(np) = nM & " monitor": cols(np) = orange
        If nA > 0 Then np = np + 1: parts(np) = nA & " avoid": cols(np) = red
        If np = 0 Then
            .Value = "-": .Font.Color = dimc
        Else
            For i = 1 To np: txt = txt & IIf(i > 1, sepTxt, "") & parts(i): Next i
            .Value = txt
            .Font.Color = dimc
            Dim pos As Long: pos = 1
            For i = 1 To np
                .Characters(pos, Len(parts(i))).Font.Color = cols(i)
                pos = pos + Len(parts(i)) + Len(sepTxt)
            Next i
        End If
    End With
    ' AVG INDEX
    With Cl(R_PASTE, 9)
        .Font.Size = 10: .Font.Bold = False: .HorizontalAlignment = xlLeft
        If nData = 0 Then
            .Value = "-": .Font.Color = dimc
        Else
            .NumberFormat = "@"
            .Value = Format$(sumIdx / nData, "+0.00;-0.00;0.00")
            .Font.Color = RGB(245, 245, 245)
            On Error Resume Next
            .Errors(xlNumberAsText).Ignore = True          ' no green "number stored as text" flag
            On Error GoTo 0
        End If
    End With
    ' REPORTED range
    With Cl(R_PASTE, 10)
        .Font.Size = 10: .Font.Bold = False: .HorizontalAlignment = xlLeft
        If dMax = 0 Then
            .Value = "-": .Font.Color = dimc
        Else
            .NumberFormat = "@"
            If dMin = dMax Then
                .Value = Format$(CDate(dMin), "mm-dd")
            Else
                .Value = Format$(CDate(dMin), "mm-dd") & " " & ChrW(&H2192) & " " & Format$(CDate(dMax), "mm-dd")
            End If
            .Font.Color = RGB(245, 245, 245)
        End If
    End With
End Sub

Private Sub TwoPart(ByVal c As Range, ByVal a As String, ByVal ca As Long, ByVal b As String, ByVal cb As Long, _
                    ByVal sep As String, ByVal cs As Long)
    With c
        .NumberFormat = "@"
        .Value = a & sep & b
        .Font.Size = 10: .Font.Bold = False: .HorizontalAlignment = xlLeft
        .Font.Color = cs
        .Characters(1, Len(a)).Font.Color = ca
        .Characters(Len(a) + Len(sep) + 1, Len(b)).Font.Color = cb
    End With
End Sub

Private Sub DrawTiles()
    Dim i As Long, cls As Long, v As String, rng As Range
    Dim sel As Long: sel = CurItem()
    ' headline
    With Rg(R_SC, 1, NUNIT - 1)
        If c_selRow > 0 Then .Cells(1, 1).Value = CellStr(c_v(c_selRow, ET_HEAD)) Else .Cells(1, 1).Value = ""
    End With
    For i = 1 To 7
        v = ""
        If c_selRow > 0 Then v = CellStr(c_v(c_selRow, ET_V1 + i - 1))
        If v = "-" Then v = ""
        cls = VClass(v)
        Set rng = Cl(R_TILE, i)
        With rng
            .Borders.LineStyle = xlNone
            .Interior.Color = ClsBg(cls, True)
            .Font.Color = ClsFg(cls)
            .Cells(1, 1).NumberFormat = VFmt(cls)
            If v = "" Then .Cells(1, 1).Value = "-" Else .Cells(1, 1).Value = v
        End With
        Call Edge(rng, xlEdgeLeft, ClsBar(cls), xlThick)
        If i = sel Then
            Call Edge(rng, xlEdgeTop, RGB(205, 205, 205), xlThin)
            Call Edge(rng, xlEdgeBottom, RGB(205, 205, 205), xlThin)
            Call Edge(rng, xlEdgeRight, RGB(205, 205, 205), xlThin)
        End If
    Next i
End Sub

Private Sub DrawDetail()
    Dim i As Long, cls As Long, v As String
    Dim sel As Long: sel = CurItem()
    With Cl(R_DHEAD, 0)
        .NumberFormat = "@"
        .Value = c_tk
    End With
    For i = 1 To 7
        v = ""
        If c_selRow > 0 Then v = CellStr(c_v(c_selRow, ET_V1 + i - 1))
        If v = "-" Then v = ""
        With Cl(R_DHEAD, i)
            .Borders.LineStyle = xlNone
            .Interior.Color = RGB(0, 0, 0)
            .Font.Bold = False
            .Font.Size = 8
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            If i = sel Then
                cls = VClass(v)
                If v = "" Then
                    .NumberFormat = "@"
                    .Value = i & " " & ItemName(i)
                    .Font.Color = RR4_ACCENT
                Else
                    .NumberFormat = VFmt(cls)
                    .Value = v
                    .Font.Color = ClsFg(cls)
                    .Interior.Color = ClsBg(cls, True)
                End If
                .Font.Bold = True
                Call Edge(Cl(R_DHEAD, i), xlEdgeLeft, IIf(v = "", RR4_ACCENT, ClsBar(cls)), xlThick)
                Call Edge(Cl(R_DHEAD, i), xlEdgeTop, RGB(205, 205, 205), xlThin)
                Call Edge(Cl(R_DHEAD, i), xlEdgeBottom, RGB(205, 205, 205), xlThin)
                Call Edge(Cl(R_DHEAD, i), xlEdgeRight, RGB(205, 205, 205), xlThin)
            Else
                .NumberFormat = "@"
                .Value = i & " " & ItemName(i)
                .Font.Color = RGB(105, 105, 105)
                Call Edge(Cl(R_DHEAD, i), xlEdgeTop, RR4_LINE, xlThin)     ' the separator above the strip
            End If
        End With
    Next i
    Dim k As Long
    For k = 1 To 4
        Dim pr As Long: pr = Array(R_PRINT, R_READ, R_QUOTE, R_SRC)(k - 1)
        If c_selRow > 0 Then
            Rg(pr, 1, NUNIT - 1).Cells(1, 1).Value = CellStr(c_v(c_selRow, DetCol(sel, k)))
        Else
            Rg(pr, 1, NUNIT - 1).Cells(1, 1).Value = ""
        End If
    Next k
End Sub

' ================================================================
'  Peer map (shapes) + peer table (cells)
' ================================================================
Private Sub DrawMapAndTable()
    Call DeleteE2Shapes(m_ws)
    Call DrawTable
    Call DrawMap
End Sub

Private Sub DrawTable()
    Dim firstR As Long: firstR = R_TFIRST + m_off
    Dim lastR As Long: lastR = m_ws.Cells(m_ws.Rows.Count, m_lc + 1).End(xlUp).Row
    If lastR < firstR + 60 Then lastR = firstR + 60
    Dim area As Range
    Set area = m_ws.Range(m_ws.Cells(firstR, 1), m_ws.Cells(lastR + 2, m_lc + NPHYS + 1))
    area.ClearContents
    area.Borders.LineStyle = xlNone
    area.Interior.Color = RGB(0, 0, 0)
    area.Font.Color = RGB(221, 221, 221)
    area.Font.Bold = False
    area.Font.Size = 9
    area.NumberFormat = "General"
    area.IndentLevel = 0
    area.HorizontalAlignment = xlGeneral
    m_ws.Range(m_ws.Rows(firstR), m_ws.Rows(lastR + 2)).RowHeight = 17

    Dim k As Long, i As Long, pr As Long, it As Long, cls As Long, b As String
    For k = 1 To p_n
        i = p_ord(k)
        pr = R_TFIRST + k - 1
        Dim selRow As Boolean: selRow = p_sel(i)
        Call Edge(Rg(pr, 0, NUNIT - 1), xlEdgeBottom, RGB(0, 0, 0), xlMedium)      ' ~2 px seam between rows
        With Cl(pr, 0)
            .NumberFormat = "@"
            .Value = p_tk(i)
            .Font.Bold = True
            .Font.Color = IIf(selRow, RR4_ACCENT, RGB(240, 240, 240))
            On Error Resume Next
            .Errors(xlNumberAsText).Ignore = True
            On Error GoTo 0
        End With
        If selRow Then
            Rg(pr, 0, NUNIT - 1).Interior.Color = RGB(26, 26, 26)
            If m_lc >= 1 Then m_ws.Cells(pr + m_off, 1).Interior.Color = RR4_ACCENT
        End If
        If Not p_has(i) Then
            With Cl(pr, 1)
                .Value = "not yet run"
                .Font.Color = RGB(105, 105, 105)
                .IndentLevel = 1
            End With
        Else
            For it = 1 To 7
                cls = VClass(p_v(i, it))
                With Cl(pr, it)
                    .Interior.Color = ClsBg(cls, False)
                    .Font.Color = ClsFg(cls)
                    .Font.Size = 8
                    .HorizontalAlignment = xlLeft
                    .IndentLevel = 1
                    .NumberFormat = VFmt(cls)
                    If p_v(i, it) = "" Then .Value = "-" Else .Value = p_v(i, it)
                End With
                Call Edge(Cl(pr, it), xlEdgeLeft, ClsBar(cls), xlMedium)
            Next it
            With Cl(pr, 8)
                .NumberFormat = "+0.00;-0.00;0.00"
                .Value = Round(p_idx(i), 4)
                .HorizontalAlignment = xlRight
                .Font.Color = RGB(235, 235, 235)
            End With
            b = BucketOf(p_idx(i))
            With Cl(pr, 9)
                .Value = b
                .Font.Bold = True
                .Font.Color = BucketColor(b)
                .HorizontalAlignment = xlLeft
                .IndentLevel = 1
            End With
            With Cl(pr, 10)
                .NumberFormat = "yyyy-mm-dd"
                .Value2 = Int(NumOr0(c_v(p_row(i), ET_REP)))
                .HorizontalAlignment = xlLeft
                .Font.Color = RGB(225, 225, 225)
            End With
        End If
    Next k
End Sub

Private Sub StyleLabel(ByVal shp As Object, ByVal txt As String, ByVal fnt As Double, ByVal clr As Long, _
                       ByVal bold As Boolean, ByVal align As Long)
    With shp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Placement = xlMove
        With .TextFrame2
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            .VerticalAnchor = 3
            With .TextRange
                .Text = txt
                .Font.Name = FONT_FACE
                .Font.Size = fnt
                .Font.Bold = IIf(bold, msoTrue, msoFalse)
                .Font.Fill.ForeColor.RGB = clr
                .ParagraphFormat.Alignment = align
            End With
        End With
    End With
End Sub

Private Sub AddText(ByVal nm As String, ByVal txt As String, ByVal cx As Double, ByVal cy As Double, ByVal w As Double, _
                    ByVal fnt As Double, ByVal clr As Long, ByVal bold As Boolean, ByVal align As Long)
    Dim h As Double: h = fnt + 5
    Dim l As Double
    If align = 1 Then l = cx Else l = cx - w / 2
    Dim s As Object: Set s = m_ws.Shapes.AddShape(1, l, cy - h / 2, w, h)
    s.Name = nm
    Call StyleLabel(s, txt, fnt, clr, bold, align)
End Sub

Private Sub AddDot(ByVal nm As String, ByVal cx As Double, ByVal cy As Double, ByVal dia As Double, ByVal clr As Long)
    Dim s As Object: Set s = m_ws.Shapes.AddShape(9, cx - dia / 2, cy - dia / 2, dia, dia)
    s.Name = nm
    s.Fill.ForeColor.RGB = clr
    s.Line.Visible = msoFalse
    s.Placement = xlMove
End Sub

Private Sub AddLine(ByVal nm As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                    ByVal clr As Long, ByVal wt As Double)
    Dim s As Object: Set s = m_ws.Shapes.AddLine(x1, y1, x2, y2)
    s.Name = nm
    s.Line.ForeColor.RGB = clr
    s.Line.Weight = wt
    s.Placement = xlMove
End Sub

' 7 ring segments (freeform polygons, clockwise from 12 o'clock: segment 1 is
' just right of the top). cols(1..7) = segment colours.
Private Sub AddRing(ByVal nm As String, ByVal cx As Double, ByVal cy As Double, ByVal dia As Double, ByRef cols() As Long)
    Dim R As Double: R = dia / 2
    Dim r0 As Double: r0 = R * 0.62
    Dim k As Long, j As Long, steps As Long: steps = 6
    Dim gap As Double: gap = 3
    Dim pi As Double: pi = 3.14159265358979
    Dim a0 As Double, a1 As Double, a As Double, fb As Object, s As Object
    For k = 1 To 7
        a0 = (-90 + (k - 1) * 360 / 7 + gap / 2) * pi / 180
        a1 = (-90 + k * 360 / 7 - gap / 2) * pi / 180
        Set fb = m_ws.Shapes.BuildFreeform(0, cx + R * Cos(a0), cy + R * Sin(a0))
        For j = 1 To steps
            a = a0 + (a1 - a0) * j / steps
            fb.AddNodes 0, 0, cx + R * Cos(a), cy + R * Sin(a)
        Next j
        For j = steps To 0 Step -1
            a = a0 + (a1 - a0) * j / steps
            fb.AddNodes 0, 0, cx + r0 * Cos(a), cy + r0 * Sin(a)
        Next j
        fb.AddNodes 0, 0, cx + R * Cos(a0), cy + R * Sin(a0)
        Set s = fb.ConvertToShape
        s.Name = nm & "_" & k
        s.Fill.ForeColor.RGB = cols(k)
        s.Line.Visible = msoFalse
        s.Placement = xlMove
    Next k
End Sub

Private Function ShapeKey(ByVal tk As String) As String
    Dim s As String, i As Long, ch As String
    For i = 1 To Len(tk)
        ch = Mid$(tk, i, 1)
        If ch Like "[A-Za-z0-9]" Then s = s & ch Else s = s & "_"
    Next i
    ShapeKey = s
End Function

Private Sub DrawMap()
    Dim axL As Double, axR As Double, axY As Double
    axL = Cl(R_AXIS, 0).Left + 62
    axR = Cl(R_AXIS, NUNIT - 1).Left + Cl(R_AXIS, NUNIT - 1).Width - 30
    axY = Cl(R_AXIS, 0).Top + Cl(R_AXIS, 0).Height / 2

    ' axis + ticks + tick labels
    Call AddLine("E2_AXIS", axL, axY, axR, axY, RGB(90, 90, 90), 0.75)
    Dim t As Long, x As Double, lbl As String
    For t = -2 To 2
        x = axL + (t * 0.25 + 0.5) * (axR - axL)
        Call AddLine("E2_TICK" & (t + 2), x, axY - 3, x, axY + 3, RGB(90, 90, 90), 0.75)
        Select Case t
            Case -2: lbl = "-0.50"
            Case -1: lbl = "-0.25"
            Case 0: lbl = "0"
            Case 1: lbl = "+0.25"
            Case 2: lbl = "+0.50"
        End Select
        Call AddText("E2_TL" & (t + 2), lbl, x, Cl(R_TICK, 0).Top + Cl(R_TICK, 0).Height / 2, 40, 7.5, RGB(120, 120, 120), False, 2)
    Next t

    ' legend
    Dim legY As Double: legY = Cl(R_LEG, 0).Top + Cl(R_LEG, 0).Height / 2
    Dim bn As Variant, bi As Long
    bn = Array("Avoid", "Monitor", "Long")
    For bi = 0 To 2
        x = axL + 4 + bi * 72
        Call AddDot("E2_LEGD" & bi, x, legY, 5, BucketColor(CStr(bn(bi))))
        Call AddText("E2_LEGT" & bi, CStr(bn(bi)), x + 7, legY, 50, 7.5, RGB(150, 150, 150), False, 1)
    Next bi

    ' key ring (which segment is which item). Its own block: centre 32 pt below the top of
    ' the MLBL row, ring 22 pt, numbers 11.5 pt outside the ring edge (clear of it and of the
    ' row above), the "key" caption a clear 6 pt under the lowest number.
    Dim kcx As Double, kcy As Double
    kcx = Cl(R_MLBL, 0).Left + 30
    kcy = Cl(R_MLBL, 0).Top + 32
    Dim kc(1 To 7) As Long, q As Long
    For q = 1 To 7: kc(q) = RGB(110, 110, 110): Next q
    Call AddRing("E2_KEY", kcx, kcy, 22, kc)
    Dim pi As Double: pi = 3.14159265358979
    Dim ang As Double
    For q = 1 To 7
        ang = (-90 + (q - 0.5) * 360 / 7) * pi / 180
        Call AddText("E2_KEYN" & q, CStr(q), kcx + 22.5 * Cos(ang), kcy + 22.5 * Sin(ang), 8, 6.5, RGB(150, 150, 150), False, 2)
    Next q
    Call AddText("E2_KEYT", "key", kcx, kcy + 22.5 + 6.25 + 6 + 6.5, 24, 7.5, RGB(130, 130, 130), False, 2)

    ' peers with data, ascending by INDEX so overlapping rings can be spread left to right
    Dim n As Long, i As Long, j As Long, tmp As Long
    Dim ord() As Long: ReDim ord(1 To IIf(p_n > 0, p_n, 1))
    For i = 1 To p_n
        If p_has(i) Then n = n + 1: ord(n) = i
    Next i
    If n = 0 Then Exit Sub
    For i = 2 To n                                       ' insertion sort: stable, INDEX ascending
        tmp = ord(i): j = i - 1
        Do While j >= 1
            If p_idx(ord(j)) > p_idx(tmp) + 0.0000000001 Then
                ord(j + 1) = ord(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        ord(j + 1) = tmp
    Next i
    Dim xt() As Double, xr() As Double, ix As Double
    ReDim xt(1 To n): ReDim xr(1 To n)
    For i = 1 To n
        ix = p_idx(ord(i))
        If ix > 0.5 Then ix = 0.5
        If ix < -0.5 Then ix = -0.5
        xt(i) = axL + (ix + 0.5) * (axR - axL)
    Next i
    Call SpreadX(xt, xr, n, 38, axL - 6, axR + 6)

    Dim ringY As Double: ringY = Cl(R_RING1, 0).Top + Cl(R_RING1, 0).Height
    Dim lblY As Double: lblY = Cl(R_MLBL, 0).Top + Cl(R_MLBL, 0).Height / 2 - 1
    Dim it As Long, cols(1 To 7) As Long, bkt As String, key As String, lw As Double
    For i = 1 To n
        j = ord(i)
        key = ShapeKey(p_tk(j))
        For it = 1 To 7: cols(it) = RingColor(VClass(p_v(j, it))): Next it
        bkt = BucketOf(p_idx(j))
        Call AddDot("E2_DOT_" & key, xt(i), axY, 6, BucketColor(bkt))
        Call AddRing("E2_ARC_" & key, xr(i), ringY, 30, cols)
        lw = Len(p_tk(j)) * 5.6 + 9
        If p_sel(j) Then
            Dim so As Object: Set so = m_ws.Shapes.AddShape(9, xr(i) - 17.5, ringY - 17.5, 35, 35)
            so.Name = "E2_SEL_" & key
            so.Fill.Visible = msoFalse
            so.Line.ForeColor.RGB = RR4_ACCENT
            so.Line.Weight = 1.75
            so.Placement = xlMove
            Dim bx As Object: Set bx = m_ws.Shapes.AddShape(5, xr(i) - lw / 2, lblY - 7.5, lw, 15)
            bx.Name = "E2_LBL_" & key
            bx.Fill.ForeColor.RGB = RR4_ACCENT
            bx.Line.Visible = msoFalse
            bx.Placement = xlMove
            With bx.TextFrame2
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                .WordWrap = msoFalse: .VerticalAnchor = 3
                .TextRange.Text = p_tk(j)
                .TextRange.Font.Name = FONT_FACE
                .TextRange.Font.Size = 7.5
                .TextRange.Font.Bold = msoTrue
                .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                .TextRange.ParagraphFormat.Alignment = 2
            End With
        Else
            Call AddText("E2_LBL_" & key, p_tk(j), xr(i), lblY, lw, 7.5, RR4_ACCENT, True, 2)
        End If
    Next i
End Sub

' 1-D label spreading: xt ascending true positions -> xr with at least `gap`
' between neighbours; a cluster is centred on the mean of its true positions.
Private Sub SpreadX(ByRef xt() As Double, ByRef xr() As Double, ByVal n As Long, ByVal gap As Double, _
                    ByVal lo As Double, ByVal hi As Double)
    Dim cs() As Long, cc() As Long, sm() As Double, nc As Long
    ReDim cs(1 To n): ReDim cc(1 To n): ReDim sm(1 To n)
    Dim i As Long
    For i = 1 To n
        nc = nc + 1
        cs(nc) = i: cc(nc) = 1: sm(nc) = xt(i)
        Do While nc > 1
            Dim pe As Double, ns As Double
            pe = sm(nc - 1) / cc(nc - 1) + cc(nc - 1) * gap / 2
            ns = sm(nc) / cc(nc) - cc(nc) * gap / 2
            If ns >= pe - 0.001 Then Exit Do
            cc(nc - 1) = cc(nc - 1) + cc(nc)
            sm(nc - 1) = sm(nc - 1) + sm(nc)
            nc = nc - 1
        Loop
    Next i
    Dim c As Long, k As Long, st As Double
    For c = 1 To nc
        st = sm(c) / cc(c) - (cc(c) - 1) * gap / 2
        If st < lo Then st = lo
        If st + (cc(c) - 1) * gap > hi Then st = hi - (cc(c) - 1) * gap
        For k = 0 To cc(c) - 1
            xr(cs(c) + k) = st + k * gap
        Next k
    Next c
End Sub

' ================================================================
'  Segment auto-fill (typed ticker)
' ================================================================
Private Sub AutoFillSegment(ByVal tk As String)
    Dim mkt As String: mkt = MktOf(tk)
    m_imLoaded = False
    Dim key As String: key = AutoSegKey(mkt, tk)
    If key = "" Then
        Cl(R_SEG, 1).Value = ""
    Else
        Cl(R_SEG, 1).Value = SegDispOf(mkt, key)
    End If
End Sub

' ================================================================
'  Events (called from the sheet code written by FinishSheetCode)
' ================================================================
Public Sub Earn2Change(ByVal ws As Worksheet, ByVal Target As Range)
    If Not NavHasRows(ws) Then Exit Sub
    If Target.CountLarge > 1 Then
        Dim mc As Variant: mc = Target.MergeCells
        If IsNull(mc) Then Exit Sub
        If Not mc Then Exit Sub
    End If
    Call Sync(ws)
    Dim t As Range: Set t = Target.Cells(1, 1)
    Dim pr As Long: pr = t.Row - m_off
    Dim u As Long: u = UnitOfCol(t.Column - m_lc - 1)
    If pr < 1 Or u < 0 Then Exit Sub                        ' nav bar rows / blank column A / spacer columns
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fail

    If pr = R_TICKER And u = 1 Then
        Call TickerChanged(ws, False)
    ElseIf pr = R_SEG And (u >= 1 And u <= 3) Then
        Call LoadCtx(ws)
        Call DrawContent(False)
    ElseIf pr = R_SEG And u = 5 Then
        Call LoadCtx(ws)
        Call DrawContent(False)
    ElseIf pr = R_TILE And u >= 1 And u <= 7 Then
        Call TileChanged(ws, u, t)
    ElseIf pr = R_SC And u >= 1 Then
        Call WriteField(ws, ET_HEAD, CellStr(Rg(R_SC, 1, NUNIT - 1).Cells(1, 1).Value))
    ElseIf (pr = R_PRINT Or pr = R_READ Or pr = R_QUOTE Or pr = R_SRC) And u >= 1 Then
        Dim k As Long
        Select Case pr
            Case R_PRINT: k = 1
            Case R_READ: k = 2
            Case R_QUOTE: k = 3
            Case R_SRC: k = 4
        End Select
        Call WriteField(ws, DetCol(CurItem(), k), CStr(Rg(pr, 1, NUNIT - 1).Cells(1, 1).Value))
    End If
    GoTo Fin
Fail:
    Dim msg As String: msg = Err.Description
    Call NavNotify("PEER EARNINGS error: " & msg, True)
Fin:
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
End Sub

' Writes one field of the (ticker, quarter) row, creating it on the first write.
Private Sub WriteField(ByVal ws As Worksheet, ByVal col As Long, ByVal txt As String)
    Call LoadCtx(ws)
    If c_tk = "" Then
        Call NavNotify("Type a TICKER first", True)
        Exit Sub
    End If
    Dim rep As Double: rep = c_rep
    If rep = 0 Then
        rep = Int(CDbl(Date))
        Cl(R_SEG, 5).NumberFormat = "yyyy-mm-dd"
        Cl(R_SEG, 5).Value2 = rep
    End If
    Dim r As Long: r = EnsureRowFor(c_tk, rep)
    If r = 0 Then Exit Sub
    Call SetField(r, col, txt)
    Call WriteDateList(c_tk)
End Sub

Private Sub TileChanged(ByVal ws As Worksheet, ByVal item As Long, ByVal t As Range)
    Dim raw As String: raw = CellStr(t.Value)
    If raw = "-" Then raw = ""
    ' canonical spelling from the item's vocabulary (any case accepted, other text kept as typed)
    Dim voc As Variant, w As Variant
    voc = Split(ItemVocab(item), ",")
    For Each w In voc
        If StrComp(CStr(w), raw, vbTextCompare) = 0 Then raw = CStr(w): Exit For
    Next w
    Call SetItem(item)
    Call LoadCtx(ws)
    If c_tk = "" Then
        Call NavNotify("Type a TICKER first", True)
        Call DrawTiles
        Exit Sub
    End If
    Call WriteField(ws, ET_V1 + item - 1, raw)
    Call LoadCtx(ws)
    Call DrawContent(False)
End Sub

' TICKER typed (keepSeg False) or picked from the peer table (keepSeg True).
Private Sub TickerChanged(ByVal ws As Worksheet, ByVal keepSeg As Boolean)
    Call Sync(ws)
    Dim tk As String: tk = UCase$(CellStr(Cl(R_TICKER, 1).Value))
    Cl(R_TICKER, 1).NumberFormat = "@"
    Cl(R_TICKER, 1).Value = tk
    Dim mkt As String: mkt = MktOf(tk)
    m_imLoaded = False
    Call WriteSegList(mkt)
    If tk = "" Then
        Cl(R_SEG, 1).Value = ""
        Cl(R_SEG, 5).Value = ""
        Cl(R_TICKER, 3).Value = ""
        Call WriteDateList("")
    Else
        Dim curKey As String
        If keepSeg Then curKey = SegKeyOf(mkt, CellStr(Cl(R_SEG, 1).Value))
        If curKey <> "" And SegHolds(mkt, curKey, tk) Then
            ' keep the segment
        Else
            Dim ak As String: ak = AutoSegKey(mkt, tk)
            If ak = "" Then Cl(R_SEG, 1).Value = "" Else Cl(R_SEG, 1).Value = SegDispOf(mkt, ak)
            If ak = "" And mkt = "US" And IsEmpty(m_im) Then Call NavNotify("IndustryMap is empty - run IMAP to get US segments", True)
        End If
        Call LoadTable
        Dim lr As Long: lr = LatestRow(BareKey(tk))
        If lr > 0 Then
            Cl(R_SEG, 5).NumberFormat = "yyyy-mm-dd"
            Cl(R_SEG, 5).Value2 = Int(NumOr0(c_v(lr, ET_REP)))
        Else
            Cl(R_SEG, 5).Value = ""
        End If
        Call WriteDateList(tk)
        Dim nm As String
        On Error Resume Next
        Dim arg As String: arg = tk
        nm = GetCompanyName(arg)
        On Error GoTo 0
        If nm = arg Or nm = tk Then nm = ""
        Cl(R_TICKER, 3).Value = nm
    End If
    Call LoadCtx(ws)
    Call DrawContent(False)
End Sub

' Selection = click: tiles / tabs pick the item, a peer's ticker becomes the TICKER, the 1 jumps to E.
Public Sub Earn2Select(ByVal ws As Worksheet, ByVal Target As Range)
    If Not NavHasRows(ws) Then Exit Sub
    If Target.CountLarge > 1 Then
        Dim mc As Variant: mc = Target.MergeCells
        If IsNull(mc) Then Exit Sub
        If Not mc Then Exit Sub
    End If
    Call Sync(ws)
    Dim t As Range: Set t = Target.Cells(1, 1)
    Dim pr As Long: pr = t.Row - m_off
    Dim u As Long: u = UnitOfCol(t.Column - m_lc - 1)
    If pr < 1 Or u < 0 Then Exit Sub
    Dim item As Long, doIt As Boolean
    If pr = R_TILE And u >= 1 And u <= 7 Then item = u: doIt = True
    If pr = R_DHEAD And u >= 1 And u <= 7 Then item = u: doIt = True
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Dim prevScr As Boolean: prevScr = Application.ScreenUpdating
    If doIt Then
        If item = CurItem() Then Exit Sub
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        On Error GoTo Fin
        Call SetItem(item)
        Call LoadCtx(ws)
        Call DrawTiles
        Call DrawDetail
        GoTo Fin
    End If
    If pr = R_LAYER And u = 1 Then
        Call LoadCtx(ws)
        Application.EnableEvents = False
        On Error GoTo Fin
        Dim mk As String: mk = IIf(c_mkt = "TW", "TW", IIf(c_mkt = "US", "US", ""))
        If c_tk = "" Then
            Call NavGoto("E", ws)
        Else
            Call modEarnings.ShowEarnings(c_tk, mk, True)
        End If
        GoTo Fin
    End If
    If pr >= R_TFIRST And u = 0 Then
        Dim pk As String: pk = UCase$(CellStr(t.Value))
        If pk = "" Then Exit Sub
        Call LoadCtx(ws)
        If BareKey(pk) = BareKey(c_tk) Then Exit Sub
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        On Error GoTo Fin
        Cl(R_TICKER, 1).NumberFormat = "@"
        Cl(R_TICKER, 1).Value = pk
        Call TickerChanged(ws, True)
        GoTo Fin
    End If
    Exit Sub
Fin:
    If Err.Number <> 0 Then Call NavNotify("PEER EARNINGS error: " & Err.Description, True)
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
End Sub

' EarnData edits (manual entry / paste): tidy the row, redraw the page's dynamic parts.
Public Sub Earn2DataChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(E2_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.Range) Is Nothing Then Exit Sub
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    Dim body As Range: Set body = Intersect(Target, lo.DataBodyRange)
    If Not body Is Nothing Then
        Dim tb As Range: Set tb = Intersect(Target, lo.ListColumns(ET_TICKER).DataBodyRange)
        If Not tb Is Nothing Then
            Dim c As Range
            For Each c In tb.Cells
                If CellStr(c.Value) <> "" Then
                    c.NumberFormat = "@"
                    c.Value = UCase$(CellStr(c.Value))
                End If
            Next c
        End If
    End If
    Dim wp As Worksheet: Set wp = ThisWorkbook.Worksheets(E2_PAGE)
    If Not wp Is Nothing Then
        If NavHasRows(wp) And NavPageCode(wp) = "E2" Then
            Call LoadCtx(wp)
            Call WriteDateList(c_tk)
            Call DrawContent(False)
        End If
    End If
    On Error GoTo 0
    Application.EnableEvents = prevEv
End Sub

' ================================================================
'  Sheet event code (document modules are not exported as .bas)
' ================================================================
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
    Call NavNotify(ws.Name & " built, but its sheet event code could not be written (enable Trust access to the VBA project object model, or paste RR4/SheetEarnings2_Code.txt by hand)", True)
End Sub
