Attribute VB_Name = "CompanyResearchSEC"
Option Explicit

' ============================================================================
'  CompanyResearchSEC  --  RR4 "Company research" sheet, lower half
' ----------------------------------------------------------------------------
'  Renders a single-ticker financial deep-dive under the MARKET SCANNER output
'  (Sanner.RunCompanyResearch). US tickers -> SEC EDGAR XBRL via
'  shared-vba/modSECData; TW co_ids -> MOPS iXBRL via shared-vba/modMOPSData.
'
'  Trigger: the "Company research" sheet's Worksheet_Change on B34
'  (see RR4/SheetCompanyResearch_Code.txt). D34 optionally overrides the
'  auto-detected market ("US" / "TW" / blank = AUTO). F34 shows status.
'
'  Requires the shared-vba modules imported into this workbook:
'    modHttp  modJsonUtil  modPrices  modSECData  modMOPSData
'  (sync with scripts/sync-shared-vba.ps1). No dependency on modCharts /
'  modTheme -- this module paints its own dark cells.
'
'  Pure ASCII (RR4 .bas convention). This module is ADDITIVE: it does not
'  touch Sanner or any existing RR4 module.
' ============================================================================

Public Const CR_SHEET       As String = "Company research"
Public Const CR_INPUT_CELL  As String = "B34"
Public Const CR_MARKET_CELL As String = "D34"
Public Const CR_STATUS_CELL As String = "F34"

Private Const TITLE_ROW      As Long = 36
Private Const HDR_ROW        As Long = 37
Private Const FIRST_DATA_ROW As Long = 38
Private Const CLEAR_LAST_ROW As Long = 400
Private Const LAST_COL       As Long = 24

Private Const WANT_ANNUAL    As Long = 4
Private Const WANT_QUARTER   As Long = 8
Private Const MONTHLY_MONTHS As Long = 24
Private Const MAX_PERIODS    As Long = 8

Private Const CLR_BG    As Long = 0             ' black -- the lower band is ALWAYS black, never reset to white
Private Const CLR_TEXT  As Long = 15132390      ' RGB(230,230,230)
Private Const CLR_HEAD  As Long = 49151         ' RGB(255,191,0) amber
Private Const CLR_MUTED As Long = 8355711       ' RGB(127,127,127)
Private Const CLR_FLAG  As Long = 3129855       ' RGB(255,192,0)-ish note
Private Const FONT_FACE As String = "Calibri"   ' whole lower band, incl. CJK entity names
Private Const COL_A_WIDTH As Double = 20
Private Const CLR_INPUT_BG As Long = 65535       ' bright yellow -- the B34 "type here" cell

' Set per run by RunDeepDive. Money values are divided by mDivisor and shown in
' billions; mUnitLbl goes into the table titles.
Private mDivisor As Double
Private mUnitLbl As String


' ---------------------------------------------------------------------------
'  Entry point
' ---------------------------------------------------------------------------
Public Sub RunDeepDive(ByVal rawTicker As String, Optional ByVal marketOverride As String = "")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CR_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & CR_SHEET & "' not found.", vbExclamation
        Exit Sub
    End If

    Call StyleInputCell(ws)

    Dim tk As String
    tk = UCase$(Trim$(rawTicker))

    Call ClearLowerBand(ws)

    If tk = "" Then
        Call SetStatus(ws, "")
        Exit Sub
    End If

    Dim mkt As String
    mkt = DetectMarket(tk, marketOverride)

    ' Both feeds report in whole currency units (SEC companyfacts = actual USD;
    ' MOPS t164sb01 iXBRL = actual NTD, ParseIxFacts already applied the scale
    ' attribute) -> divide by 1e9 for billions. The monthly-revenue feed
    ' (t21sc03 HTML) is the exception: it is in thousands, handled in RenderMonthly.
    mDivisor = 1000000000#
    If mkt = "TW" Then mUnitLbl = "NTD bn" Else mUnitLbl = "USD bn"

    Application.ScreenUpdating = False
    Call SetStatus(ws, "Fetching " & tk & " (" & mkt & ") ...")

    On Error GoTo Fail
    If mkt = "TW" Then
        Call RenderTw(ws, tk)
    Else
        Call RenderUs(ws, tk)
    End If

    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    Call SetStatus(ws, "Error: " & Err.Description)
End Sub


' ---------------------------------------------------------------------------
'  US  (SEC EDGAR / XBRL)
' ---------------------------------------------------------------------------
Private Sub RenderUs(ByVal ws As Worksheet, ByVal ticker As String)
    Dim r As Object
    Set r = modSECData.GetXbrlFinancials(ticker, WANT_ANNUAL, WANT_QUARTER)

    If Not r("ok") Then
        Call SetStatus(ws, r("error"))
        Exit Sub
    End If

    ' periods: quarterly filings, newest first (as returned), capped
    Dim fq As Collection
    Set fq = r("filings10Q")
    Dim nP As Long
    nP = fq.Count
    If nP < 1 Then
        Call SetStatus(ws, "No 10-Q filings for " & ticker & " (only annual on record).")
        Exit Sub
    End If
    If nP > MAX_PERIODS Then nP = MAX_PERIODS

    Dim labels() As String
    ReDim labels(1 To nP)
    Dim accns() As String, rdates() As String, forms() As String
    ReDim accns(1 To nP): ReDim rdates(1 To nP): ReDim forms(1 To nP)

    Dim i As Long
    For i = 1 To nP
        Dim f As Object
        Set f = fq(i)
        accns(i) = CStr(f("accn"))
        rdates(i) = CStr(f("reportDate"))
        forms(i) = CStr(f("form"))
        Dim fyfp As String
        fyfp = modSECData.XbrlFyFp(r("allMaps"), accns(i), rdates(i), forms(i))
        If fyfp <> "" And fyfp <> "|" Then
            Dim parts() As String
            parts = Split(fyfp, "|")
            labels(i) = "FY" & parts(0) & " " & parts(1)
        Else
            labels(i) = rdates(i)
        End If
    Next i

    Dim keys As Variant, disp As Variant
    keys = MetricRowKeys()
    disp = MetricRowLabels()

    Dim vals As Object
    Set vals = CreateObject("Scripting.Dictionary")
    Dim ki As Long
    For ki = LBound(keys) To UBound(keys)
        Dim arr() As Variant
        ReDim arr(1 To nP)
        For i = 1 To nP
            arr(i) = ToNum(modSECData.XbrlValue(r("maps"), CStr(keys(ki)), accns(i), rdates(i), forms(i)))
        Next i
        vals(CStr(keys(ki))) = arr
    Next ki

    Dim startRow As Long
    startRow = RenderTable(ws, TITLE_ROW, _
        CStr(r("entityName")) & "  (" & ticker & ")  -  SEC EDGAR XBRL, quarterly (" & mUnitLbl & ")", _
        labels, keys, disp, vals, _
        "Source: SEC EDGAR companyfacts / submissions API, CIK " & CStr(r("cik")) & _
        ". Flow items = standalone quarter. Margins/DSO derived. " & _
        "Blank CapEx/Interest = tag not present in that filing (same as the standalone SEC tool).")

    Call SetStatus(ws, "Done. " & nP & " quarters shown.")
End Sub


' ---------------------------------------------------------------------------
'  TW  (MOPS iXBRL) -- quarterly financials + monthly revenue
' ---------------------------------------------------------------------------
Private Sub RenderTw(ByVal ws As Worksheet, ByVal coId As String)
    Dim r As Object
    Set r = modMOPSData.GetTwFinancials(coId, WANT_QUARTER)
    If Not r("ok") Then
        Call SetStatus(ws, r("error"))
        Exit Sub
    End If

    Dim rows As Collection
    Set rows = r("quarterlyRows")
    Dim nP As Long
    nP = rows.Count
    If nP < 1 Then
        Call SetStatus(ws, "No MOPS quarterly data for " & coId & ".")
        Exit Sub
    End If
    If nP > MAX_PERIODS Then nP = MAX_PERIODS

    Dim labels() As String
    ReDim labels(1 To nP)
    Dim metricsByCol() As Object
    ReDim metricsByCol(1 To nP)
    Dim i As Long
    For i = 1 To nP
        Dim spec As Object
        Set spec = rows(i)                       ' newest first (as accumulated)
        labels(i) = CStr(spec("fyLabel")) & " " & CStr(spec("fpLabel"))
        Set metricsByCol(i) = spec("metrics")
    Next i

    Dim keys As Variant, disp As Variant
    keys = MetricRowKeys()
    disp = MetricRowLabels()

    Dim vals As Object
    Set vals = CreateObject("Scripting.Dictionary")
    Dim ki As Long
    For ki = LBound(keys) To UBound(keys)
        Dim arr() As Variant
        ReDim arr(1 To nP)
        For i = 1 To nP
            arr(i) = ToNum(metricsByCol(i)(CStr(keys(ki))))
        Next i
        vals(CStr(keys(ki))) = arr
    Next ki

    Dim afterFin As Long
    afterFin = RenderTable(ws, TITLE_ROW, _
        CStr(r("entityName")) & "  (" & coId & ")  -  MOPS consolidated, quarterly (" & mUnitLbl & ")", _
        labels, keys, disp, vals, _
        "Source: MOPS t164sb01 inline-XBRL (REPORT_ID=C, fallback A). " & _
        "Shares/tax rate/DSO derived. No per-share dividend tag on this feed.")

    ' ---- monthly revenue ----
    Dim mr As Object
    Set mr = modMOPSData.GetTwMonthlyRevenue(coId, MONTHLY_MONTHS)
    If Not mr("ok") Then
        Call WriteFlag(ws, afterFin + 1, "Monthly revenue: " & mr("error"))
        Call SetStatus(ws, "Done (financials only; monthly revenue unavailable).")
        Exit Sub
    End If

    Call RenderMonthly(ws, afterFin + 2, CStr(r("entityName")), mr)
    Call SetStatus(ws, "Done.")
End Sub

Private Sub RenderMonthly(ByVal ws As Worksheet, ByVal atRow As Long, ByVal entityName As String, ByVal mr As Object)
    Dim labels As Variant, rev As Variant, yoy As Variant
    labels = mr("labels"): rev = mr("rev"): yoy = mr("yoy")
    Dim n As Long
    n = UBound(labels) - LBound(labels) + 1

    Call WriteTitle(ws, atRow, entityName & "  -  monthly revenue, last " & n & " months (" & mUnitLbl & ")")
    Dim hdr As Long
    hdr = atRow + 1

    ws.Cells(hdr, 1).Value = "Month"
    ws.Cells(hdr + 1, 1).Value = "Revenue"
    ws.Cells(hdr + 2, 1).Value = "YoY %"

    Dim c As Long, j As Long
    For j = LBound(labels) To UBound(labels)
        c = 2 + (j - LBound(labels))
        If c > LAST_COL Then Exit For
        ws.Cells(hdr, c).Value = labels(j)
        ' t21sc03 monthly revenue is in THOUSANDS of NTD -> /1e6 for billions
        If IsNumeric(rev(j)) Then ws.Cells(hdr + 1, c).Value = CDbl(rev(j)) / 1000000# Else ws.Cells(hdr + 1, c).Value = "-"
        ws.Cells(hdr + 1, c).NumberFormat = "#,##0.00"
        If IsNumeric(yoy(j)) And Not IsEmpty(yoy(j)) Then
            ws.Cells(hdr + 2, c).Value = CDbl(yoy(j))
            ws.Cells(hdr + 2, c).NumberFormat = "0.0%"
        Else
            ws.Cells(hdr + 2, c).Value = "-"
        End If
    Next j

    Call PaintBand(ws, atRow, hdr + 2)
    ws.Cells(atRow, 1).Font.Color = CLR_HEAD
    ws.Cells(atRow, 1).Font.Bold = True
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, LAST_COL)).Font.Color = CLR_HEAD
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr + 2, 1)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(hdr, 2), ws.Cells(hdr + 2, LAST_COL)).HorizontalAlignment = xlRight

    Call WriteFlag(ws, hdr + 3, "Flag: newest month may post-date the latest quarterly report -- " & _
        "read it as leading confirmation/contradiction of guidance. Source: MOPS t21sc03 monthly filing.")
End Sub


' ---------------------------------------------------------------------------
'  Shared table renderer
'  metrics down col A from FIRST_DATA_ROW; periods across cols B.. ; then a
'  derived block (margins, DSO). Returns the last row written.
' ---------------------------------------------------------------------------
Private Function RenderTable(ByVal ws As Worksheet, ByVal titleRow As Long, ByVal title As String, _
        ByRef labels() As String, ByVal keys As Variant, ByVal disp As Variant, ByVal vals As Object, _
        ByVal sourceNote As String) As Long

    Dim nP As Long
    nP = UBound(labels) - LBound(labels) + 1

    Call WriteTitle(ws, titleRow, title)

    ' header row of period labels
    Dim c As Long, i As Long
    ws.Cells(HDR_ROW, 1).Value = "Metric"
    For i = 1 To nP
        c = 1 + i
        If c > LAST_COL Then Exit For
        ws.Cells(HDR_ROW, c).Value = labels(i)
    Next i

    ' raw metric rows
    Dim rr As Long
    rr = FIRST_DATA_ROW
    Dim ki As Long
    For ki = LBound(keys) To UBound(keys)
        ws.Cells(rr, 1).Value = disp(ki)
        Dim arr As Variant
        arr = vals(CStr(keys(ki)))
        Dim isMoney As Boolean
        isMoney = IsMoneyKey(CStr(keys(ki)))
        For i = 1 To nP
            c = 1 + i
            If c > LAST_COL Then Exit For
            If IsNumeric(arr(i)) Then
                If isMoney Then
                    ws.Cells(rr, c).Value = CDbl(arr(i)) / mDivisor
                Else
                    ws.Cells(rr, c).Value = CDbl(arr(i))
                End If
                ws.Cells(rr, c).NumberFormat = NumFmtFor(CStr(keys(ki)))
            Else
                ws.Cells(rr, c).Value = "-"
            End If
        Next i
        rr = rr + 1
    Next ki

    ' derived block
    rr = rr + 1
    ws.Cells(rr, 1).Value = "-- derived --"
    ws.Cells(rr, 1).Font.Color = CLR_MUTED
    rr = rr + 1

    rr = DerivedRow(ws, rr, "Gross margin", labels, vals, "GrossProfit", "Revenue", "0.0%")
    rr = DerivedRow(ws, rr, "Operating margin", labels, vals, "OperatingIncome", "Revenue", "0.0%")
    rr = DerivedRow(ws, rr, "Net margin", labels, vals, "NetIncome", "Revenue", "0.0%")
    rr = DerivedRow(ws, rr, "FCF (CFO - CapEx)", labels, vals, "Cfo", "CapEx", "#,##0.00", True, True)
    rr = DerivedRow(ws, rr, "Interest coverage (x)", labels, vals, "OperatingIncome", "InterestExpense", "0.0")
    rr = DsoRow(ws, rr, "DSO (days)", labels, vals)

    Dim lastRow As Long
    lastRow = rr - 1

    Call PaintBand(ws, titleRow, lastRow + 1)
    ws.Cells(titleRow, 1).Font.Color = CLR_HEAD
    ws.Cells(titleRow, 1).Font.Bold = True
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW, LAST_COL)).Font.Color = CLR_HEAD
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW, LAST_COL)).Font.Bold = True

    ' tidy: labels flush left, every number flush right across the whole grid
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(lastRow, 1)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(HDR_ROW, 2), ws.Cells(lastRow, LAST_COL)).HorizontalAlignment = xlRight

    Call WriteFlag(ws, lastRow + 1, "Flag: " & sourceNote)

    RenderTable = lastRow + 1
End Function

Private Function DerivedRow(ByVal ws As Worksheet, ByVal atRow As Long, ByVal label As String, _
        ByRef labels() As String, ByVal vals As Object, ByVal numKey As String, ByVal denKey As String, _
        ByVal fmt As String, Optional ByVal subtract As Boolean = False, Optional ByVal moneyOut As Boolean = False) As Long
    ws.Cells(atRow, 1).Value = label
    Dim nP As Long, i As Long, c As Long
    nP = UBound(labels) - LBound(labels) + 1
    Dim numA As Variant, denA As Variant
    numA = vals(numKey): denA = vals(denKey)
    For i = 1 To nP
        c = 1 + i
        If c > LAST_COL Then Exit For
        Dim outV As Variant
        outV = "-"
        If IsNumeric(numA(i)) And IsNumeric(denA(i)) Then
            If subtract Then
                outV = CDbl(numA(i)) - CDbl(denA(i))
            ElseIf CDbl(denA(i)) <> 0 Then
                outV = CDbl(numA(i)) / CDbl(denA(i))
            End If
        End If
        If IsNumeric(outV) Then
            If moneyOut Then
                ws.Cells(atRow, c).Value = CDbl(outV) / mDivisor
            Else
                ws.Cells(atRow, c).Value = CDbl(outV)
            End If
            ws.Cells(atRow, c).NumberFormat = fmt
        Else
            ws.Cells(atRow, c).Value = "-"
        End If
    Next i
    DerivedRow = atRow + 1
End Function

Private Function DsoRow(ByVal ws As Worksheet, ByVal atRow As Long, ByVal label As String, _
        ByRef labels() As String, ByVal vals As Object) As Long
    ws.Cells(atRow, 1).Value = label
    Dim nP As Long, i As Long, c As Long
    nP = UBound(labels) - LBound(labels) + 1
    Dim arV As Variant, revV As Variant
    arV = vals("Ar"): revV = vals("Revenue")
    For i = 1 To nP
        c = 1 + i
        If c > LAST_COL Then Exit For
        ' VBA And is NOT short-circuit -- guard CDbl behind nested Ifs so a
        ' blank Revenue cell (e.g. a MOPS Q4 quarterly period) can't hit CDbl("").
        Dim okCell As Boolean
        okCell = False
        If IsNumeric(arV(i)) And IsNumeric(revV(i)) Then
            If CDbl(revV(i)) <> 0 Then okCell = True
        End If
        If okCell Then
            ws.Cells(atRow, c).Value = CDbl(arV(i)) / CDbl(revV(i)) * 91#
            ws.Cells(atRow, c).NumberFormat = "0.0"
        Else
            ws.Cells(atRow, c).Value = "-"
        End If
    Next i
    DsoRow = atRow + 1
End Function


' ---------------------------------------------------------------------------
'  helpers
' ---------------------------------------------------------------------------
Private Function MetricRowKeys() As Variant
    MetricRowKeys = Array("Revenue", "GrossProfit", "OperatingIncome", "NetIncome", "Eps", _
        "Cfo", "CapEx", "Assets", "Liabilities", "StockholdersEquity", "Cash", _
        "Inventory", "Ar", "CurrentAssets", "CurrentLiabilities", "InterestExpense", "EffectiveTaxRate")
End Function

Private Function MetricRowLabels() As Variant
    MetricRowLabels = Array("Revenue", "Gross profit", "Operating income", "Net income", "Diluted EPS", _
        "CFO", "CapEx", "Total assets", "Total liabilities", "Equity", "Cash & equiv.", _
        "Inventory", "Accounts receivable", "Current assets", "Current liabilities", "Interest expense", "Effective tax rate")
End Function

Private Function NumFmtFor(ByVal key As String) As String
    Select Case key
        Case "Eps":              NumFmtFor = "0.00"
        Case "EffectiveTaxRate": NumFmtFor = "0.0%"
        Case Else:               NumFmtFor = "#,##0.00"   ' billions, 2 dp
    End Select
End Function

Private Function IsMoneyKey(ByVal key As String) As Boolean
    IsMoneyKey = Not (key = "Eps" Or key = "EffectiveTaxRate")
End Function

Private Function ToNum(ByVal v As Variant) As Variant
    On Error GoTo bad
    If IsEmpty(v) Then ToNum = "": Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Then ToNum = "": Exit Function
    End If
    If IsNumeric(v) Then ToNum = CDbl(v) Else ToNum = ""
    Exit Function
bad:
    ToNum = ""
End Function

Private Function DetectMarket(ByVal tk As String, ByVal override As String) As String
    Dim o As String
    o = UCase$(Trim$(override))
    If o = "US" Or o = "TW" Then
        DetectMarket = o
        Exit Function
    End If
    ' AUTO: a TW co_id is all digits (3-6). Anything with a letter is a US ticker.
    Dim i As Long, allDigits As Boolean
    allDigits = (Len(tk) >= 3 And Len(tk) <= 6)
    For i = 1 To Len(tk)
        If InStr("0123456789", Mid$(tk, i, 1)) = 0 Then allDigits = False: Exit For
    Next i
    DetectMarket = IIf(allDigits, "TW", "US")
End Function

Private Sub StyleInputCell(ByVal ws As Worksheet)
    ' B34 = the ticker / co_id input. Yellow ground, black bold text, so it
    ' stands out on the all-black sheet as "type here". Formatting only -- does
    ' not fire Worksheet_Change.
    With ws.Range(CR_INPUT_CELL)
        .Interior.Color = CLR_INPUT_BG
        .Font.Color = 0
        .Font.Bold = True
        .Font.Name = FONT_FACE
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub ClearLowerBand(ByVal ws As Worksheet)
    ' Clear content, then repaint the whole band BLACK. Never xlNone / white.
    With ws.Range(ws.Cells(TITLE_ROW, 1), ws.Cells(CLEAR_LAST_ROW, 60))
        .Clear
        .Interior.Color = CLR_BG
        .Font.Color = CLR_TEXT
        .Font.Name = FONT_FACE
        .Font.Bold = False
        .Font.Italic = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(1).ColumnWidth = COL_A_WIDTH
End Sub

Private Sub PaintBand(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    With ws.Range(ws.Cells(r1, 1), ws.Cells(r2, LAST_COL))
        .Interior.Color = CLR_BG
        .Font.Color = CLR_TEXT
        .Font.Name = FONT_FACE
    End With
End Sub

Private Sub WriteTitle(ByVal ws As Worksheet, ByVal atRow As Long, ByVal txt As String)
    ws.Cells(atRow, 1).Value = txt
End Sub

Private Sub WriteFlag(ByVal ws As Worksheet, ByVal atRow As Long, ByVal txt As String)
    ws.Cells(atRow, 1).Value = txt
    With ws.Range(ws.Cells(atRow, 1), ws.Cells(atRow, LAST_COL))
        .Interior.Color = CLR_BG
        .Font.Color = CLR_FLAG
        .Font.Italic = True
        .Font.Name = FONT_FACE
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub SetStatus(ByVal ws As Worksheet, ByVal msg As String)
    Application.EnableEvents = False
    ws.Range(CR_STATUS_CELL).Value = msg
    Application.EnableEvents = True
    Application.StatusBar = IIf(msg = "", False, "Company research: " & msg)
End Sub
