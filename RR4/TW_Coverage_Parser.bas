Attribute VB_Name = "TW_Coverage_Parser"
Option Explicit

' ═══════════════════════════════════════════════════════════════════
'  My-TW-Coverage GitHub Parser  (完整合併版 + 全 Bug 修正)
'  Repo : Timeverse/My-TW-Coverage
'
'  功能 A (單筆) : 在「設定」B3 輸入代號 → 執行 FetchSingleReport
'  功能 B (批次) : 執行 ParsePilotReports → 寫入四張工作表
' ═══════════════════════════════════════════════════════════════════

Private Const RAW_BASE      As String = "https://raw.githubusercontent.com/Timeverse/My-TW-Coverage/master/"
Private Const API_BASE      As String = "https://api.github.com/repos/Timeverse/My-TW-Coverage/"

' ── 工作表名稱 ────────────────────────────────────────────────────
Private Const SETTINGS_SHEET As String = "Settings"
Private Const SH_REPORT     As String = "Company Report"   ' 單筆報告頁（不含非法字元）
Private Const SH_OVERVIEW   As String = "公司總覽"
Private Const SH_BIZ        As String = "業務與供應鏈"
Private Const SH_ANNUAL     As String = "年度財務"
Private Const SH_QTR        As String = "季度財務"

' ═══════════════════════════════════════════════════════════════════
'  共用輔助：安全偵測 String() 陣列是否已初始化
'  原始寫法 (Not paths) = -1 對陣列做 bitwise NOT → Type Mismatch 崩潰
' ═══════════════════════════════════════════════════════════════════
Private Function IsArrayInitialized(ByRef arr() As String) As Boolean
    On Error Resume Next
    Dim lb As Long
    lb = LBound(arr)
    IsArrayInitialized = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub FetchSingleReport()
    Dim token As String: token = GetToken()
    Dim Ticker As String
    Dim wsReport As Worksheet
    
    ' 優先從「公司報告」工作表的 B1 取得代號
    EnsureSheet SH_REPORT
    Set wsReport = ThisWorkbook.Sheets(SH_REPORT)
    Ticker = Trim(wsReport.Range("B1").Value)
    
    ' 如果報告頁 B1 是空的，則回頭去找「設定」工作表的 B3
    If Ticker = "" Then
        Ticker = Trim(ThisWorkbook.Sheets(SETTINGS_SHEET).Range("B3").Value)
    End If

    If Ticker = "" Then
        MsgBox "錯誤：請在「公司報告」B1 或「設定」B3 輸入股票代號。", vbCritical
        Exit Sub
    End If

    Application.StatusBar = "讀取 GitHub 檔案清單..."
    Dim paths() As String
    paths = GetAllMDPaths(token)

    If Not IsArrayInitialized(paths) Then
        MsgBox "錯誤：無法取得目錄，請檢查 Token。", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    ' 搜尋路徑... (其餘邏輯不變)
    Dim targetPath As String: targetPath = ""
    Dim i As Long
    For i = 0 To UBound(paths)
        If InStr(paths(i), Ticker) > 0 Then
            targetPath = paths(i)
            Exit For
        End If
    Next i

    If targetPath = "" Then
        MsgBox "搜尋結束：找不到包含「" & Ticker & "」的檔案。", vbInformation
        Exit Sub
    End If

    Dim mdContent As String
    mdContent = FetchRaw(targetPath, token)
    If Len(mdContent) > 0 Then
        FormatReportSheet mdContent, targetPath
    End If
    Application.StatusBar = False
End Sub
' ═══════════════════════════════════════════════════════════════════
'  功能 C：產業搜尋
'  在「設定」B4 輸入產業關鍵字（如 Semiconductors）→ 執行 SearchBySector
'  結果寫入「產業搜尋」工作表，深色主題
' ═══════════════════════════════════════════════════════════════════
Sub SearchBySector()
    Const SH_SECTOR As String = "產業搜尋"
    Dim token As String: token = GetToken()
    Dim keyword As String
    keyword = Trim(ThisWorkbook.Sheets(SETTINGS_SHEET).Range("B4").Value)

    If keyword = "" Then
        MsgBox "請在「設定」B4 輸入產業關鍵字（如 Semiconductors）", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "讀取清單..."

    Dim paths() As String: paths = GetAllMDPaths(token)
    If Not IsArrayInitialized(paths) Then
        MsgBox "無法取得清單，請確認 Token。", vbCritical
        GoTo CleanupSector
    End If

    ' 收集符合路徑（不分大小寫比對產業資料夾名稱）
    Dim hits() As String: ReDim hits(0 To UBound(paths))
    Dim hitCount As Long
    Dim i As Long
    For i = 0 To UBound(paths)
        If InStr(LCase(paths(i)), LCase(keyword)) > 0 Then
            hits(hitCount) = paths(i): hitCount = hitCount + 1
        End If
    Next i

    If hitCount = 0 Then
        MsgBox "找不到包含「" & keyword & "」的資料夾。", vbInformation
        GoTo CleanupSector
    End If

    ' 建立 / 清空工作表
    EnsureSheet SH_SECTOR
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_SECTOR)
    ws.cells.Clear
    ApplySheetDarkBg ws, hitCount + 10

    ' ── 頁首統計列 ───────────────────────────────────────────────
    With ws.cells(1, 1)
        .Value = "SECTOR SCAN  |  " & UCase(keyword) & "  |  " & hitCount & " COMPANIES  |  " & Now()
        .Font.Bold = True: .Font.Size = 11
        .Font.Color = RGB(13, 13, 13)
        .Interior.Color = RGB(212, 175, 55)
    End With
    ws.Range(ws.cells(1, 1), ws.cells(1, 10)).Merge
    ws.Rows(1).RowHeight = 24

    ' ── 欄位標題 ──────────────────────────────────────────────────
    Dim hd As Variant
    hd = Array("TICKER", "COMPANY", "SECTOR", "INDUSTRY", "MKT CAP", "EV", "P/E TTM", "P/B", "EV/EBITDA", "PATH")
    Dim ci As Long
    For ci = 0 To UBound(hd)
        With ws.cells(2, ci + 1)
            .Value = hd(ci)
            .Font.Bold = True
            .Font.Color = RGB(212, 175, 55)
            .Interior.Color = RGB(30, 30, 30)
            .HorizontalAlignment = xlCenter
        End With
    Next ci
    ws.Rows(2).RowHeight = 20

    ' ── 逐一抓取並解析 ───────────────────────────────────────────
    Dim r As Long: r = 3
    Dim j As Long
    For j = 0 To hitCount - 1
        Application.StatusBar = "抓取 " & (j + 1) & " / " & hitCount & ":  " & hits(j)
        Dim mc As String: mc = FetchRaw(hits(j), token)
        If Len(mc) = 0 Then GoTo NextHit

        Dim fp() As String: fp = Split(hits(j), "/")
        Dim fnb As String: fnb = Replace(fp(UBound(fp)), ".md", "")
        Dim tk As String, co As String
        If InStr(fnb, "_") > 0 Then
            tk = Split(fnb, "_")(0): co = Split(fnb, "_")(1)
        Else
            tk = fnb: co = fnb
        End If

        ' 估值指標
        Dim fSec As String: fSec = ExtractSectionLoose(mc, "財務概況")
        Dim vSec As String: vSec = ExtractSubSectionLoose(fSec, "估值指標")
        Dim pe As String, pb As String, evebitda As String
        pe = ParseValuationField(vSec, 0)
        pb = ParseValuationField(vSec, 3)
        evebitda = ParseValuationField(vSec, 4)

        ' 寫入每一欄
        Dim rowVals(0 To 9) As Variant
        rowVals(0) = tk
        rowVals(1) = co
        rowVals(2) = RegexFirst(mc, "\*\*板塊[:：]\*\*\s*(.+)")
        rowVals(3) = RegexFirst(mc, "\*\*產業[:：]\*\*\s*(.+)")
        rowVals(4) = RegexFirst(mc, "\*\*市值[:：]\*\*\s*(.+)")
        rowVals(5) = RegexFirst(mc, "\*\*企業價值[:：]\*\*\s*(.+)")
        rowVals(6) = pe
        rowVals(7) = pb
        rowVals(8) = evebitda
        rowVals(9) = hits(j)

        Dim k As Long
        For k = 0 To 9
            With ws.cells(r, k + 1)
                .Font.Color = RGB(210, 210, 210)
                Select Case k
                    Case 0  ' Ticker
                        .Value = rowVals(k)
                        .Font.Bold = True
                        .Font.Color = RGB(212, 175, 55)
                    Case 6, 7, 8  ' 估值數字
                        If IsNumeric(rowVals(k)) Then
                            .Value = CDbl(rowVals(k))
                            .NumberFormat = "0.00"
                            .Font.Color = RGB(24, 144, 255)
                        Else
                            .Value = rowVals(k)
                            .Font.Color = RGB(130, 130, 130)
                        End If
                    Case Else
                        .Value = rowVals(k)
                End Select
                ' 偶數列微亮底色
                If (r Mod 2) = 0 Then
                    .Interior.Color = RGB(22, 22, 22)
                End If
            End With
        Next k
        r = r + 1
NextHit:
        If j Mod 10 = 0 Then DoEvents
    Next j

    ' 欄寬
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G:I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 50

    ws.Activate
    ws.cells(1, 1).Select
    ActiveWindow.FreezePanes = False
    ws.cells(3, 1).Select
    ActiveWindow.FreezePanes = True

CleanupSector:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If hitCount > 0 Then
        MsgBox "完成！共找到 " & hitCount & " 家 " & keyword & " 公司。", vbInformation
    End If
End Sub

' ── 從估值表格取出指定欄（資料列第一列 = 標題列下面那列）────────
Private Function ParseValuationField(vSec As String, colIdx As Long) As String
    If vSec = "" Then Exit Function
    Dim normalized As String
    normalized = Replace(Replace(vSec, vbCrLf, vbLf), vbCr, vbLf)
    Dim lines() As String: lines = Split(normalized, vbLf)
    Dim dataRowCount As Long
    Dim ln As Variant
    For Each ln In lines
        Dim ls As String: ls = Trim(CStr(ln))
        If Left(ls, 1) = "|" And InStr(ls, "---") = 0 And Len(ls) > 1 Then
            dataRowCount = dataRowCount + 1
            If dataRowCount = 2 Then  ' 第 2 個有效列 = 數值列
                Dim cA() As String: cA = SplitCells(ls)
                If colIdx <= UBound(cA) Then ParseValuationField = Trim(cA(colIdx))
                Exit Function
            End If
        End If
    Next ln
End Function

' ── 基礎背景渲染：全域純黑 (無死角版) ──
Private Sub ApplySheetDarkBg(ws As Worksheet, lastRow As Long)
    ' 1. 將整張工作表的所有儲存格均設為純黑底色與無邊框
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Borders.LineStyle = xlNone
    End With
    
    ' 2. 僅針對會有資料出現的範圍設定預設字型，避免消耗多餘記憶體
    With ws.Range(ws.cells(1, 1), ws.cells(lastRow + 30, 22))
        .Font.Color = RGB(210, 210, 210)
        .Font.Name = "Consolas"
        .Font.Size = 10
    End With
    
    ws.Tab.Color = RGB(21, 96, 130)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub
' ── 單筆報告排版主引擎 ──
Private Sub FormatReportSheet(content As String, mdPath As String)
    EnsureSheet SH_REPORT
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_REPORT)
    
    ' 從路徑解析 Ticker 與公司名稱
    Dim parts() As String: parts = Split(mdPath, "/")
    Dim filename As String: filename = Replace(parts(UBound(parts)), ".md", "")
    Dim Ticker As String, company As String
    If InStr(filename, "_") > 0 Then
        Ticker = Split(filename, "_")(0)
        company = Split(filename, "_")(1)
    Else
        Ticker = filename: company = ""
    End If

    ' 1. 清空並初始化純黑背景
    ws.cells.Clear
    ApplySheetDarkBg ws, 120

    ' 2. 建立頂部輸入控制區
    With ws.cells(1, 1)
        .Value = "查詢代號:"
        .Font.Color = RGB(130, 130, 130)
        .Font.Bold = True: .Font.Size = 9
    End With
    
    With ws.cells(1, 2)
        .Value = Ticker
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0) ' 純黑背景
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(212, 175, 55)
    End With
    
    With ws.cells(1, 3)
        .Value = " ← 輸入後直接執行"
        .Font.Color = RGB(100, 100, 100): .Font.Size = 8
    End With
    ws.Rows(1).RowHeight = 25

    Dim nr As Long
    ' Row 2: 公司名稱大標題
    With ws.cells(2, 1)
        .Value = company & " " & Ticker
        .Font.Size = 16: .Font.Bold = True
        .Font.Color = RGB(212, 175, 55)
        .Font.Name = "Consolas"
        .Interior.Color = RGB(0, 0, 0)
    End With
    ws.Rows(2).RowHeight = 35
    
    nr = 3
    
    ' ── Row 3: 基本資訊 stat bar ──
    Dim statLabels As Variant, statVals As Variant
    statLabels = Array("板塊", "產業", "市值", "企業價值")
    statVals = Array( _
        RegexFirst(content, "\*\*板塊[:：]\*\*\s*(.+)"), _
        RegexFirst(content, "\*\*產業[:：]\*\*\s*(.+)"), _
        RegexFirst(content, "\*\*市值[:：]\*\*\s*(.+)"), _
        RegexFirst(content, "\*\*企業價值[:：]\*\*\s*(.+)") _
    )
    Dim si As Long
    For si = 0 To 3
        Dim col As Long: col = si * 3 + 1
        With ws.cells(nr, col)
            .Value = UCase(statLabels(si))
            .Font.Color = RGB(130, 130, 130)
            .Font.Size = 8: .Font.Bold = True
        End With
        With ws.cells(nr, col + 1)
            .Value = statVals(si)
            .Font.Color = RGB(212, 175, 55)
            .Font.Bold = True: .Font.Size = 10
        End With
    Next si
    ws.Rows(nr).RowHeight = 18
    nr = nr + 1

    Dim sepRange As Range: Set sepRange = ws.Range(ws.cells(nr, 1), ws.cells(nr, 18))
    sepRange.Interior.Color = RGB(212, 175, 55)
    ws.Rows(nr).RowHeight = 2
    nr = nr + 1

    ' ── 業務簡介 ──
    DarkSectionHeader ws, nr, "  BUSINESS OVERVIEW", RGB(0, 0, 0), RGB(212, 175, 55)
    nr = nr + 1
    Dim biz As String
    biz = CleanWikilinks(StripMetaLines(ExtractSectionLoose(content, "業務")))
    With ws.cells(nr, 1)
        .Value = biz
        .WrapText = True
        .Font.Color = RGB(190, 190, 190)
        .Font.Size = 10
    End With
    ws.Rows(nr).RowHeight = 80
    ws.Range(ws.cells(nr, 1), ws.cells(nr, 18)).Merge
    nr = nr + 1

    ' ── 供應鏈 (已優化 Regex 邏輯) ──
    DarkSectionHeader ws, nr, "  SUPPLY CHAIN", RGB(0, 0, 0), RGB(24, 144, 255)
    nr = nr + 1
    Dim scSection As String: scSection = ExtractSectionLoose(content, "供應鏈")
    Dim scData(0 To 2) As Variant
    ' 使用高度寬容的正則匹配：無視星號位置，抓取冒號後所有文字
    scData(0) = Array("UPSTREAM", CleanWikilinks(RegexFirst(scSection, "上游.*?[:：][\*\s]*([^\r\n]+)")))
    scData(1) = Array("MIDSTREAM", CleanWikilinks(RegexFirst(scSection, "中游.*?[:：][\*\s]*([^\r\n]+)")))
    scData(2) = Array("DOWNSTREAM", CleanWikilinks(RegexFirst(scSection, "下游.*?[:：][\*\s]*([^\r\n]+)")))
    Dim sci As Long
    For sci = 0 To 2
        With ws.cells(nr, 1)
            .Value = scData(sci)(0)
            .Font.Color = RGB(130, 130, 130): .Font.Bold = True: .Font.Size = 8
        End With
        With ws.cells(nr, 2)
            .Value = scData(sci)(1)
            .Font.Color = RGB(210, 210, 210)
            .WrapText = True
        End With
        ws.Range(ws.cells(nr, 2), ws.cells(nr, 12)).Merge
        ws.Rows(nr).RowHeight = 18
        nr = nr + 1
    Next sci

    ' ── 主要客戶 / 供應商 ──
    DarkSectionHeader ws, nr, "  CUSTOMERS  &  SUPPLIERS", RGB(0, 0, 0), RGB(24, 144, 255)
    nr = nr + 1
    Dim csSection As String: csSection = ExtractSectionLoose(content, "主要客戶")
    Dim csData(0 To 1) As Variant
    csData(0) = Array("CUSTOMERS", CleanWikilinks(ExtractListItems(ExtractSubSectionLoose(csSection, "主要客戶"))))
    csData(1) = Array("SUPPLIERS", CleanWikilinks(ExtractListItems(ExtractSubSectionLoose(csSection, "主要供應商"))))
    Dim csi As Long
    For csi = 0 To 1
        With ws.cells(nr, 1)
            .Value = csData(csi)(0)
            .Font.Color = RGB(130, 130, 130): .Font.Bold = True: .Font.Size = 8
        End With
        With ws.cells(nr, 2)
            .Value = csData(csi)(1)
            .Font.Color = RGB(210, 210, 210)
            .WrapText = True
        End With
        ws.Range(ws.cells(nr, 2), ws.cells(nr, 18)).Merge
        ws.Rows(nr).RowHeight = 18
        nr = nr + 1
    Next csi

    ' ── 財務概況 ──
    Dim finSection As String: finSection = ExtractSectionLoose(content, "財務")

    If Len(finSection) > 0 Then
        ' ── 估值指標 (矩陣轉置直式排列) ──
        Dim vSec As String: vSec = ExtractSubSectionLoose(finSection, "估值指標")
        If Len(vSec) > 0 Then
            
            ' 【新增邏輯】提取括號內的股價基準日期
            Dim priceDateStr As String
            ' 尋找 "as of 2026-03-26" 這種格式
            priceDateStr = RegexFirst(content, "估值指標\s*\([^)]*as of\s*([0-9\-]+)[^)]*\)")
            
            Dim valHeader As String
            If priceDateStr <> "" Then
                valHeader = "  VALUATION METRICS (As of " & priceDateStr & ")"
            Else
                valHeader = "  VALUATION METRICS"
            End If

            nr = nr + 1
            ' 將抓取到的日期融合進紅色區塊標題
            DarkSectionHeader ws, nr, valHeader, RGB(0, 0, 0), RGB(255, 100, 100)
            nr = nr + 1
            
            Dim vLines() As String: vLines = Split(Replace(vSec, vbCr, ""), vbLf)
            Dim vHead() As String, vData() As String
            Dim validRowCount As Long: validRowCount = 0
            Dim lineIt As Variant
            
            ' 拆解 Markdown 表格為兩組陣列
            For Each lineIt In vLines
                Dim lineStr As String: lineStr = Trim(CStr(lineIt))
                If Left(lineStr, 1) = "|" And InStr(lineStr, "---") = 0 Then
                    validRowCount = validRowCount + 1
                    If validRowCount = 1 Then
                        vHead = SplitCells(lineStr)
                    ElseIf validRowCount = 2 Then
                        vData = SplitCells(lineStr)
                        Exit For
                    End If
                End If
            Next lineIt
            
            ' 垂直印出陣列
            If validRowCount = 2 Then
                Dim vi As Long
                For vi = 0 To UBound(vHead)
                    If Trim(vHead(vi)) <> "" Then ' 過濾掉潛在的空欄位
                        With ws.cells(nr, 1) ' 欄位名稱放在A欄
                            .Value = Trim(vHead(vi))
                            .Font.Color = RGB(130, 130, 130)
                            .Font.Bold = True
                            .HorizontalAlignment = xlRight
                        End With
                        With ws.cells(nr, 2) ' 數值放在B欄
                            .Value = Trim(SafeCell(vData, vi))
                            .Font.Color = RGB(255, 100, 100)
                            .Font.Bold = True
                            .HorizontalAlignment = xlLeft
                        End With
                        ws.Range(ws.cells(nr, 2), ws.cells(nr, 4)).Merge
                        nr = nr + 1
                    End If
                Next vi
            End If
        End If

        ' 年度財務
        Dim aSec As String: aSec = ExtractSubSectionLoose(finSection, "年度")
        If Len(aSec) > 0 Then
            nr = nr + 1
            DarkSectionHeader ws, nr, "  ANNUAL FINANCIALS", RGB(0, 0, 0), RGB(82, 196, 26)
            nr = nr + 1
            nr = DarkRenderMdTable(ws, aSec, nr, RGB(0, 0, 0), RGB(82, 196, 26))
        End If

        ' 季度財務
        Dim qSec As String: qSec = ExtractSubSectionLoose(finSection, "季度")
        If Len(qSec) > 0 Then
            nr = nr + 1
            DarkSectionHeader ws, nr, "  QUARTERLY FINANCIALS", RGB(0, 0, 0), RGB(250, 173, 20)
            nr = nr + 1
            nr = DarkRenderMdTable(ws, qSec, nr, RGB(0, 0, 0), RGB(250, 173, 20))
        End If
    End If

    ' 欄寬重設
    ws.Columns("A").ColumnWidth = 26
    Dim c As Long
    For c = 2 To 8
        ws.Columns(c).ColumnWidth = 14
    Next c

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.cells(4, 1).Select
    ActiveWindow.FreezePanes = True
    ws.Range("B1").Select
End Sub
' ── 深色區塊標題 ─────────────────────────────────────────────────
Private Sub DarkSectionHeader(ws As Worksheet, r As Long, _
                               title As String, bgColor As Long, accentColor As Long)
    Dim rng As Range: Set rng = ws.Range(ws.cells(r, 1), ws.cells(r, 18))
    rng.Interior.Color = bgColor
    ws.cells(r, 1).Interior.Color = accentColor
    ws.cells(r, 1).Value = ""  ' accent strip
    With ws.cells(r, 2)
        .Value = title
        .Font.Bold = True
        .Font.Color = accentColor
        .Font.Size = 10
        .Font.Name = "Consolas"
        .Interior.Color = bgColor
    End With
    ws.Range(ws.cells(r, 2), ws.cells(r, 18)).Merge
    ws.Rows(r).RowHeight = 22
End Sub

' ── Markdown 表格渲染器 (純黑版) ──
Private Function DarkRenderMdTable(ws As Worksheet, mdBlock As String, _
                                    startRow As Long, _
                                    headerBg As Long, accentColor As Long) As Long
    Dim normalized As String
    normalized = Replace(Replace(mdBlock, vbCrLf, vbLf), vbCr, vbLf)
    Dim lines() As String: lines = Split(normalized, vbLf)

    Dim r As Long: r = startRow
    Dim isHeader As Boolean: isHeader = True
    Dim rowNum As Long: rowNum = 0
    Dim ln As Variant

    For Each ln In lines
        Dim ls As String: ls = Trim(CStr(ln))
        If Left(ls, 1) = "|" And InStr(ls, "---") = 0 And Len(ls) > 1 Then
            Dim cA() As String: cA = SplitCells(ls)
            Dim c As Long
            For c = 0 To UBound(cA)
                Dim cv As String: cv = CleanWikilinks(Trim(cA(c)))
                With ws.cells(r, c + 1)
                    ' 移除交錯底色，全部設定為純黑
                    .Interior.Color = IIf(isHeader, headerBg, RGB(0, 0, 0))
                    
                    If isHeader Then
                        .Value = cv
                        .Font.Bold = True
                        .Font.Color = accentColor
                        .HorizontalAlignment = xlCenter
                    ElseIf c = 0 Then
                        .Value = cv
                        .Font.Color = RGB(170, 170, 170)
                        .Font.Bold = True
                    ElseIf IsNumeric(cv) Then
                        .Value = CDbl(cv)
                        .NumberFormat = "0.00"
                        Dim labelCell As String
                        labelCell = LCase(Trim(cA(0)))
                        If InStr(labelCell, "margin") > 0 Or InStr(labelCell, "income") > 0 Or _
                           InStr(labelCell, "profit") > 0 Then
                            If CDbl(cv) >= 0 Then
                                .Font.Color = RGB(82, 196, 26)
                            Else
                                .Font.Color = RGB(220, 53, 69)
                            End If
                        Else
                            .Font.Color = RGB(210, 210, 210)
                        End If
                        .HorizontalAlignment = xlRight
                    Else
                        .Value = cv
                        .Font.Color = RGB(130, 130, 130)
                        .HorizontalAlignment = xlCenter
                    End If
                End With
            Next c
            If Not isHeader Then rowNum = rowNum + 1
            isHeader = False
            r = r + 1
        End If
    Next ln
    DarkRenderMdTable = r
End Function
' ── 寬鬆版 ExtractSection（heading 後可接任意字元）──────────────
Private Function ExtractSectionLoose(content As String, heading As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.MultiLine = False
    re.pattern = "## " & heading & "[^\r\n]*\r?\n([\s\S]+?)(?=\r?\n## |$)"
    Dim m As Object: Set m = re.Execute(content)
    If m.count > 0 Then ExtractSectionLoose = m(0).SubMatches(0) Else ExtractSectionLoose = ""
End Function

' ── 寬鬆版 ExtractSubSection（### heading 後可接任意字元）────────
Private Function ExtractSubSectionLoose(content As String, heading As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.MultiLine = False
    re.pattern = "### " & heading & "[^\r\n]*\r?\n([\s\S]+?)(?=\r?\n###|$)"
    Dim m As Object: Set m = re.Execute(content)
    If m.count > 0 Then ExtractSubSectionLoose = m(0).SubMatches(0) Else ExtractSubSectionLoose = ""
End Function

' ═══════════════════════════════════════════════════════════════════
'  功能 B：批次解析所有公司
' ═══════════════════════════════════════════════════════════════════
Sub ParsePilotReports()
    Dim token As String: token = GetToken()

    If token = "" Then
        MsgBox "請先在「" & SETTINGS_SHEET & "」工作表 B1 填入 GitHub Token。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    SetupSheets

    Application.StatusBar = "讀取檔案清單..."
    Dim paths() As String
    paths = GetAllMDPaths(token)

    If Not IsArrayInitialized(paths) Then
        MsgBox "無法取得檔案清單，請確認 Token 是否有效或網路狀態。", vbCritical
        GoTo Cleanup
    End If

    Dim total As Long: total = UBound(paths) + 1
    Dim i As Long, written As Long
    Dim nextRows(1 To 4) As Long
    Dim n As Long
    For n = 1 To 4: nextRows(n) = 2: Next n

    For i = 0 To UBound(paths)
        Application.StatusBar = "解析 " & (i + 1) & " / " & total & "  (" & paths(i) & ")"
        Dim mdContent As String
        mdContent = FetchRaw(paths(i), token)
        If Len(mdContent) > 0 Then
            WriteAllSheets mdContent, paths(i), nextRows
            written = written + 1
        End If
        If (i + 1) Mod 50 = 0 Then DoEvents
    Next i

    AutoFitSheets

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "完成！共解析 " & written & " 家公司", vbInformation
End Sub

' ═══════════════════════════════════════════════════════════════════
'  工作表初始化（批次用）
' ═══════════════════════════════════════════════════════════════════
Private Sub SetupSheets()
    EnsureSheet SETTINGS_SHEET
    With ThisWorkbook.Sheets(SETTINGS_SHEET)
        If .Range("A1").Value = "" Then
            .Range("A1").Value = "GitHub Token"
            .Range("A1").Font.Bold = True
            .Range("B1").Value = "（請貼上 ghp_xxxx Token）"
            .Range("B1").Font.Color = RGB(150, 150, 150)
        End If
        If .Range("A3").Value = "" Then
            .Range("A3").Value = "單筆查詢代號"
            .Range("A3").Font.Bold = True
        End If
    End With

    EnsureSheet SH_OVERVIEW
    EnsureSheet SH_BIZ
    EnsureSheet SH_ANNUAL
    EnsureSheet SH_QTR

    ClearDataSheet SH_OVERVIEW
    ClearDataSheet SH_BIZ
    ClearDataSheet SH_ANNUAL
    ClearDataSheet SH_QTR

    WriteHeaders
End Sub

Private Sub EnsureSheet(ByVal shName As String)
    Dim safeName As String: safeName = CleanSheetName(shName)
    If Len(safeName) = 0 Then Exit Sub

    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(safeName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        On Error Resume Next
        ws.Name = safeName
        On Error GoTo 0
    End If
End Sub

Private Function CleanSheetName(ByVal txt As String) As String
    Dim illegal As Variant: illegal = Array("\", "/", "?", "*", "[", "]", ":")
    Dim i As Integer
    txt = Replace(txt, vbCr, ""): txt = Replace(txt, vbLf, "")
    For i = LBound(illegal) To UBound(illegal)
        txt = Replace(txt, illegal(i), "_")
    Next i
    If Len(txt) > 31 Then txt = Left(txt, 31)
    CleanSheetName = Trim(txt)
End Function

Private Sub ClearDataSheet(ByVal shName As String)
    With ThisWorkbook.Sheets(shName)
        If .UsedRange.Rows.count > 1 Then
            .Rows("2:" & .UsedRange.Rows.count + 1).ClearContents
        End If
    End With
End Sub

Private Sub WriteHeaders()
    WriteHeaderRow SH_OVERVIEW, Array("Ticker", "公司名稱", "產業資料夾", "板塊", "產業", "市值（百萬台幣）", "企業價值（百萬台幣）"), RGB(31, 78, 121)
    WriteHeaderRow SH_BIZ, Array("Ticker", "公司名稱", "產業資料夾", "業務簡介", "上游", "中游", "下游", "主要客戶", "主要供應商"), RGB(55, 86, 35)
    WriteHeaderRow SH_ANNUAL, Array("Ticker", "公司名稱", "產業資料夾", "年度營收1", "年度營收2", "年度營收3", "毛利率1", "毛利率2", "毛利率3", "營業利益率1", "營業利益率2", "營業利益率3", "EPS1", "EPS2", "EPS3"), RGB(123, 44, 44)
    WriteHeaderRow SH_QTR, Array("Ticker", "公司名稱", "產業資料夾", "Q1營收", "Q2營收", "Q3營收", "Q4營收", "Q1 EPS", "Q2 EPS", "Q3 EPS", "Q4 EPS"), RGB(123, 82, 19)
End Sub

Private Sub WriteHeaderRow(shName As String, headers As Variant, bgColor As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(shName)
    Dim i As Long
    For i = 0 To UBound(headers)
        With ws.cells(1, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Font.Name = "Arial"
            .Interior.Color = bgColor
            .HorizontalAlignment = xlCenter
        End With
    Next i
    ws.Rows(1).RowHeight = 24
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.cells(2, 4).Select
    ActiveWindow.FreezePanes = True
End Sub

' ═══════════════════════════════════════════════════════════════════
'  API 請求
' ═══════════════════════════════════════════════════════════════════
Private Function GetAllMDPaths(token As String) As String()
    Dim url As String: url = API_BASE & "git/trees/master?recursive=1"
    Dim json As String: json = HttpGet(url, token)
    If json = "" Then Exit Function

    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.pattern = """path"":""(Pilot_Reports/[^/]+/[^""]+\.md)"""

    Dim matches As Object: Set matches = re.Execute(json)
    If matches.count = 0 Then Exit Function

    Dim result() As String
    ReDim result(0 To matches.count - 1)
    Dim i As Long
    For i = 0 To matches.count - 1
        result(i) = matches(i).SubMatches(0)
    Next i
    GetAllMDPaths = result
End Function

' ── FetchRaw：改用 Contents API + raw accept header ──────────────
'  原本走 raw.githubusercontent.com，對中文檔名 URL 編碼敏感，容易靜默失敗。
'  Contents API 讓 GitHub 伺服器端解析路徑，Accept: v3.raw 直接回傳純文字，
'  兩者合用可完全繞過中文編碼問題，且失敗時能看到真實 HTTP 狀態碼。
Private Function FetchRaw(mdPath As String, token As String) As String
    ' 使用 Contents API（不是 raw 域名）
    Dim url As String: url = API_BASE & "contents/" & UrlEncodePath(mdPath)

    On Error GoTo ErrHandler
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Excel-VBA-TW-Coverage"
    If token <> "" Then http.setRequestHeader "Authorization", "token " & token
    ' 關鍵：告訴 GitHub API 直接回傳原始文字，而非 JSON 包裝
    http.setRequestHeader "Accept", "application/vnd.github.v3.raw"
    http.send

    If http.status = 200 Then
        FetchRaw = http.responseText
    Else
        ' 顯示真實錯誤碼，方便診斷（404=找不到路徑, 401=Token 無效, 403=權限不足）
        MsgBox "下載失敗！" & vbCrLf & _
               "HTTP " & http.status & "：" & http.StatusText & vbCrLf & vbCrLf & _
               "路徑：" & mdPath & vbCrLf & _
               "請求網址：" & url, vbCritical, "FetchRaw 診斷"
        FetchRaw = ""
    End If
    Exit Function
ErrHandler:
    MsgBox "網路連線失敗：" & Err.Description, vbCritical, "FetchRaw 網路錯誤"
    FetchRaw = ""
End Function

' HttpGet 僅供 GetAllMDPaths（抓目錄清單）使用，維持 JSON accept header
Private Function HttpGet(url As String, token As String) As String
    On Error GoTo ErrHandler
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Excel-VBA-TW-Coverage"
    If token <> "" Then http.setRequestHeader "Authorization", "token " & token
    http.setRequestHeader "Accept", "application/vnd.github.v3+json"
    http.send
    If http.status = 200 Then HttpGet = http.responseText Else HttpGet = ""
    Exit Function
ErrHandler:
    HttpGet = ""
End Function

' ═══════════════════════════════════════════════════════════════════
'  批次資料解析與寫入
' ═══════════════════════════════════════════════════════════════════
Private Sub WriteAllSheets(content As String, mdPath As String, nextRows() As Long)
    Dim parts() As String: parts = Split(mdPath, "/")
    If UBound(parts) < 2 Then Exit Sub

    Dim sectorFolder As String: sectorFolder = parts(1)
    Dim filename As String:     filename = parts(2)
    Dim fnBase As String:       fnBase = Left(filename, Len(filename) - 3)

    Dim Ticker As String, company As String
    Dim underPos As Long: underPos = InStr(fnBase, "_")
    If underPos > 0 Then
        Ticker = Left(fnBase, underPos - 1)
        company = Mid(fnBase, underPos + 1)
    Else
        Ticker = fnBase: company = fnBase
    End If

    Dim 板塊 As String, 產業 As String, 市值 As String, EV As String, 業務 As String
    Dim 上游 As String, 中游 As String, 下游 As String, 客戶 As String, 供應商 As String
    Dim ann(1 To 12) As String
    Dim qtr(1 To 8) As String

    板塊 = RegexFirst(content, "\*\*板塊[:：]\*\*\s*(.+)")
    產業 = RegexFirst(content, "\*\*產業[:：]\*\*\s*(.+)")
    市值 = Replace(RegexFirst(content, "\*\*市值[:：]\*\*\s*([\d,，]+)"), ",", "")
    EV = Replace(RegexFirst(content, "\*\*企業價值[:：]\*\*\s*([\d,，]+)"), ",", "")

    業務 = CleanWikilinks(StripMetaLines(ExtractSection(content, "業務簡介")))
    If Len(業務) > 400 Then 業務 = Left(業務, 400) & "..."

    Dim scSection As String: scSection = ExtractSection(content, "供應鏈位置")
    上游 = CleanWikilinks(RegexFirst(scSection, "\*\*上游[:：]\*\*\s*(.+)"))
    中游 = CleanWikilinks(RegexFirst(scSection, "\*\*中游[:：]\*\*\s*(.+)"))
    下游 = CleanWikilinks(RegexFirst(scSection, "\*\*下游[:：]\*\*\s*(.+)"))

    Dim csSection As String: csSection = ExtractSection(content, "主要客戶及供應商")
    客戶 = CleanWikilinks(ExtractListItems(ExtractSubSection(csSection, "主要客戶")))
    供應商 = CleanWikilinks(ExtractListItems(ExtractSubSection(csSection, "主要供應商")))

    ParseFinancials ExtractSection(content, "財務概況"), ann, qtr

    Dim ws As Worksheet, c As Long

    Set ws = ThisWorkbook.Sheets(SH_OVERVIEW)
    ws.cells(nextRows(1), 1).Resize(1, 7).Value = Array(Ticker, company, sectorFolder, 板塊, 產業, 市值, EV)

    Set ws = ThisWorkbook.Sheets(SH_BIZ)
    ws.cells(nextRows(2), 1).Resize(1, 9).Value = Array(Ticker, company, sectorFolder, 業務, 上游, 中游, 下游, 客戶, 供應商)

    Set ws = ThisWorkbook.Sheets(SH_ANNUAL)
    ws.cells(nextRows(3), 1).Resize(1, 3).Value = Array(Ticker, company, sectorFolder)
    For c = 1 To 12: ws.cells(nextRows(3), 3 + c) = ann(c): Next c

    Set ws = ThisWorkbook.Sheets(SH_QTR)
    ws.cells(nextRows(4), 1).Resize(1, 3).Value = Array(Ticker, company, sectorFolder)
    For c = 1 To 8: ws.cells(nextRows(4), 3 + c) = qtr(c): Next c

    Dim n As Long
    For n = 1 To 4: nextRows(n) = nextRows(n) + 1: Next n
End Sub

Private Sub ParseFinancials(finSection As String, ann() As String, qtr() As String)
    If finSection = "" Then Exit Sub

    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.pattern = "(\|.+\|(\r?\n|\r)(\| *[-:| ]+ *\|(\r?\n|\r))?(\|.+\|(\r?\n|\r))+)"

    Dim tables As Object: Set tables = re.Execute(finSection)
    Dim t As Object
    For Each t In tables
        Dim lines() As String: lines = Split(t.Value, vbLf)
        Dim rowLines As New Collection
        Dim ln As Variant
        For Each ln In lines
            Dim lineStr As String
            lineStr = Trim(Replace(CStr(ln), vbCr, ""))
            If Left(lineStr, 1) = "|" And InStr(lineStr, "---") = 0 Then
                rowLines.Add SplitCells(lineStr)
            End If
        Next ln

        Dim row As Variant
        For Each row In rowLines
            Dim cells() As String: cells = row
            If UBound(cells) < 1 Then GoTo NextRow
            Dim label As String: label = Trim(cells(0))
            Dim v1 As String: v1 = SafeCell(cells, 1)
            Dim v2 As String: v2 = SafeCell(cells, 2)
            Dim v3 As String: v3 = SafeCell(cells, 3)
            Dim v4 As String: v4 = SafeCell(cells, 4)

            If InStr(label, "營收") > 0 Or InStr(label, "Revenue") > 0 Then
                If v4 <> "" Then
                    qtr(1) = v1: qtr(2) = v2: qtr(3) = v3: qtr(4) = v4
                Else
                    ann(1) = v1: ann(2) = v2: ann(3) = v3
                End If
            ElseIf InStr(label, "毛利率") > 0 Or InStr(label, "Gross") > 0 Then
                ann(4) = v1: ann(5) = v2: ann(6) = v3
            ElseIf InStr(label, "營業利益率") > 0 Or InStr(label, "Operating") > 0 Then
                ann(7) = v1: ann(8) = v2: ann(9) = v3
            ElseIf InStr(label, "EPS") > 0 Then
                If v4 <> "" Then
                    qtr(5) = v1: qtr(6) = v2: qtr(7) = v3: qtr(8) = v4
                Else
                    ann(10) = v1: ann(11) = v2: ann(12) = v3
                End If
            End If
NextRow:
        Next row
    Next t
End Sub

' ═══════════════════════════════════════════════════════════════════
'  Markdown 解析工具
' ═══════════════════════════════════════════════════════════════════

' [BUG FIX] \Z 在 VBScript RegExp 不支援，最後一個段落永遠抓不到。
' MultiLine=False 時，$ 等同整個字串結尾，行為正確。
Private Function ExtractSection(content As String, heading As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.MultiLine = False
    re.pattern = "## " & heading & "\r?\n([\s\S]+?)(?=\r?\n## |$)"
    Dim m As Object: Set m = re.Execute(content)
    If m.count > 0 Then ExtractSection = m(0).SubMatches(0) Else ExtractSection = ""
End Function

' [BUG FIX] 同上
Private Function ExtractSubSection(content As String, heading As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.MultiLine = False
    re.pattern = "### " & heading & "\r?\n([\s\S]+?)(?=\r?\n###|$)"
    Dim m As Object: Set m = re.Execute(content)
    If m.count > 0 Then ExtractSubSection = m(0).SubMatches(0) Else ExtractSubSection = ""
End Function

Private Function StripMetaLines(text As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.MultiLine = True
    re.pattern = "^\*\*.+?[:：]\*\*.+\r?\n?"
    StripMetaLines = Trim(re.Replace(text, ""))
End Function

Private Function CleanWikilinks(text As String) As String
    If text = "" Then Exit Function
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.pattern = "\[\[(.+?)\]\]"
    CleanWikilinks = re.Replace(text, "$1")
End Function

Private Function ExtractListItems(text As String) As String
    If text = "" Then Exit Function
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.MultiLine = True
    re.pattern = "^[-*]\s*(.+)"
    Dim matches As Object: Set matches = re.Execute(text)
    If matches.count = 0 Then Exit Function
    Dim results() As String: ReDim results(0 To matches.count - 1)
    Dim i As Long
    For i = 0 To matches.count - 1
        results(i) = Trim(matches(i).SubMatches(0))
    Next i
    ExtractListItems = Join(results, "、")
End Function

Private Function RegexFirst(text As String, pattern As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.MultiLine = True
    re.pattern = pattern
    Dim m As Object: Set m = re.Execute(text)
    If m.count > 0 Then RegexFirst = Trim(m(0).SubMatches(0)) Else RegexFirst = ""
End Function

Private Function SafeCell(cells() As String, idx As Long) As String
    If idx <= UBound(cells) Then SafeCell = Trim(Replace(cells(idx), ",", "")) Else SafeCell = ""
End Function

Private Function SplitCells(line As String) As String()
    Dim inner As String: inner = line
    If Left(inner, 1) = "|" Then inner = Mid(inner, 2)
    If Right(inner, 1) = "|" Then inner = Left(inner, Len(inner) - 1)
    Dim parts() As String: parts = Split(inner, "|")
    Dim i As Long
    For i = 0 To UBound(parts): parts(i) = Trim(parts(i)): Next i
    SplitCells = parts
End Function

' ═══════════════════════════════════════════════════════════════════
'  URL 編碼
' ═══════════════════════════════════════════════════════════════════

' [BUG FIX] Hex(code) 對 code<16 只回傳 1 位，輸出 %A 而非 %0A。
' Right("0" & Hex(code), 2) 強制補零至兩位。
Private Function UrlEncodePath(path As String) As String
    Dim result As String, c As String
    Dim i As Long, code As Long
    For i = 1 To Len(path)
        c = Mid(path, i, 1)
        code = AscW(c)
        ' [BUG FIX] AscW 回傳有號 16-bit，U+8000~U+FFFF 的字元（如「電」U+96FB=38651）
        ' 超過 32767 後 AscW 回傳負數（-26885），負數 < &H800 為 True，
        ' 誤用 2-byte 公式產生錯誤位元組序列 → GitHub 回 404。
        ' 加 65536 將負數還原為正確的 Unicode code point。
        If code < 0 Then code = code + 65536
        If (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or _
           (code >= 48 And code <= 57) Or c = "/" Or c = "_" Or c = "-" Or c = "." Then
            result = result & c
        ElseIf code < 128 Then
            result = result & "%" & Right("0" & Hex(code), 2)
        Else
            result = result & EncodeUTF8Char(c)
        End If
    Next i
    UrlEncodePath = result
End Function

Private Function EncodeUTF8Char(c As String) As String
    Dim code As Long: code = AscW(c)
    If code < 0 Then code = code + 65536  '[BUG FIX] AscW 對 U+8000+ 回傳負數，加 65536 還原正確 code point
    Dim result As String
    If code < &H800 Then
        result = "%" & Right("0" & Hex(&HC0 Or (code \ 64)), 2) & _
                 "%" & Right("0" & Hex(&H80 Or (code Mod 64)), 2)
    ElseIf code < &H10000 Then
        result = "%" & Right("0" & Hex(&HE0 Or (code \ 4096)), 2) & _
                 "%" & Right("0" & Hex(&H80 Or ((code \ 64) Mod 64)), 2) & _
                 "%" & Right("0" & Hex(&H80 Or (code Mod 64)), 2)
    End If
    EncodeUTF8Char = UCase(result)
End Function

Private Sub AutoFitSheets()
    Dim shNames As Variant: shNames = Array(SH_OVERVIEW, SH_BIZ, SH_ANNUAL, SH_QTR)
    Dim s As Variant
    For Each s In shNames
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CStr(s))
        ws.Columns.AutoFit
        If CStr(s) = SH_BIZ Then
            If ws.Columns(4).ColumnWidth > 60 Then ws.Columns(4).ColumnWidth = 60
            ws.Columns(4).WrapText = True
        End If
    Next s
End Sub

' ═══════════════════════════════════════════════════════════════════
'  Token 管理
' ═══════════════════════════════════════════════════════════════════
Private Function GetToken() As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    If ws Is Nothing Then Exit Function
    Dim val As String: val = Trim(ws.Range("B1").Value)
    If val <> "" And InStr(val, "請貼上") = 0 Then GetToken = val
    On Error GoTo 0
End Function

Sub SetToken()
    Dim token As String
    token = InputBox("請貼上 GitHub Personal Access Token：", "設定 Token")
    If token = "" Then Exit Sub
    EnsureSheet SETTINGS_SHEET
    With ThisWorkbook.Sheets(SETTINGS_SHEET)
        .Range("A1").Value = "GitHub Token"
        .Range("A1").Font.Bold = True
        .Range("B1").Value = token
        .Range("B1").Font.Color = RGB(100, 100, 100)
        .Range("B1").NumberFormat = "@"
    End With
    MsgBox "Token 已儲存到「" & SETTINGS_SHEET & "」B1。", vbInformation
End Sub


