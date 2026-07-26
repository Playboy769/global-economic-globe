Attribute VB_Name = "CorelationMatrix"
Sub BuildCorrelationMatrix()
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("Correlation")
    On Error GoTo 0
    If wsC Is Nothing Then
        Set wsC = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsC.Name = "Correlation"
    End If

    ' ¢w¢w Layout constants ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    Dim MATRIX_ROW As Long: MATRIX_ROW = 2
    Dim MATRIX_COL As Long: MATRIX_COL = 4

    ' ¢w¢w Read days BEFORE clearing (preserve user input) ¢w¢w¢w¢w¢w¢w¢w¢w¢w
    Dim nDays As Long: nDays = 60
    If IsNumeric(wsC.cells(2, 2).Value) And wsC.cells(2, 2).Value > 0 Then
        nDays = CLng(wsC.cells(2, 2).Value)
    End If

    ' ¢w¢w Base style ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    wsC.cells.Clear
    With wsC.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 8
    End With
    wsC.Activate
    ActiveWindow.DisplayGridlines = False

    ' ¢w¢w Title ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    With wsC.cells(1, 1)
        .Value = "US SECTOR CORRELATION MATRIX"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 13
    End With

    ' ¢w¢w Days label + input ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    wsC.cells(2, 1).Value = "DAYS:"
    wsC.cells(2, 1).Font.Color = RGB(150, 150, 150)
    wsC.cells(2, 1).Font.Bold = True
    With wsC.cells(2, 2)
        .Value = nDays
        .Interior.Color = RGB(30, 30, 0)
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
    End With

    ' ¢w¢w Sector ETF list ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    Dim tickers As Variant
    Dim labels  As Variant
    tickers = Array( _
        "XLK", "XLF", "XLV", "XLE", "XLY", "XLP", "XLI", "XLU", "XLRE", "XLB", "XLC", _
        "IGV", "KBE", "XBI", "JETS", "XHB", "XRT", "XME", "XOP", "KIE", _
        "SMH", "SOXX", "ARKK", "ARKG", "KWEB", "GDX", "SLX", "MOO", "PHO", "ICLN")
    labels = Array( _
        "Tech", "Finance", "Health", "Energy", "ConDisc", "ConStap", "Industrl", "Utility", "RealEst", "Material", "CommSvc", _
        "SoftETF", "Banks", "BioTech", "Airlines", "HomeBld", "Retail", "Metals", "OilGas", "Insure", _
        "SemiETF", "SOXX", "ARKK", "ARKG", "China", "Gold", "Steel", "Agri", "Water", "CleanE")

    Dim nSec As Integer: nSec = UBound(tickers) + 1

    ' ¢w¢w Fetch return data ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    Application.StatusBar = "Fetching price history for " & nSec & " sectors..."
    Dim retData() As Double
    ReDim retData(0 To nSec - 1, 0 To nDays - 2)

    Dim si As Integer
    For si = 0 To nSec - 1
        Application.StatusBar = "Fetching [" & si + 1 & "/" & nSec & "] " & tickers(si)
        DoEvents
        Dim prices() As Double
        prices = FetchPriceArray(CStr(tickers(si)), nDays)
        Dim j As Long
        For j = 0 To nDays - 2
            If j + 1 <= UBound(prices) And prices(j) > 0 Then
                retData(si, j) = (prices(j + 1) - prices(j)) / prices(j)
            Else
                retData(si, j) = 0
            End If
        Next j
    Next si

    ' ¢w¢w Column widths ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    wsC.Columns(1).ColumnWidth = 36
    wsC.Columns(2).ColumnWidth = 12
    wsC.Columns(3).Hidden = True
    Dim ci As Integer
    For ci = 0 To nSec - 1
        wsC.Columns(MATRIX_COL + ci).ColumnWidth = 6
    Next ci

    ' ¢w¢w Row heights ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    wsC.Rows(1).RowHeight = 22
    wsC.Rows(MATRIX_ROW).RowHeight = 18
    Dim ri As Integer
    For ri = 0 To nSec - 1
        wsC.Rows(MATRIX_ROW + 1 + ri).RowHeight = 15
    Next ri

    ' ¢w¢w Column headers (rotated) ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    For ci = 0 To nSec - 2
        With wsC.cells(MATRIX_ROW, MATRIX_COL + ci)
            .Value = labels(ci)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 8
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .Orientation = 0
            .WrapText = False
        End With
    Next ci

    ' ¢w¢w Matrix ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    For ri = 0 To nSec - 1
        Dim dataRow As Long: dataRow = MATRIX_ROW + 1 + ri

        With wsC.cells(dataRow, 2)
            .Value = labels(ri)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 8
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With

        For ci = 0 To nSec - 1
            Dim corr As Double
            If ri = ci Then
                corr = 1
            Else
                corr = CalcCorrelation(retData, ri, ci, nDays - 1)
            End If

            With wsC.cells(dataRow, MATRIX_COL + ci)
                .Value = corr
                .NumberFormat = "0.00"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 8
                .Font.Bold = (ri = ci)

                Select Case True
                    Case corr >= 0.7
                        .Interior.Color = RGB(120, 0, 0)
                        .Font.Color = RGB(255, 180, 180)
                    Case corr >= 0.4
                        .Interior.Color = RGB(70, 15, 15)
                        .Font.Color = RGB(230, 130, 130)
                    Case corr >= 0.1
                        .Interior.Color = RGB(25, 25, 25)
                        .Font.Color = RGB(190, 190, 190)
                    Case corr >= -0.1
                        .Interior.Color = RGB(10, 10, 10)
                        .Font.Color = RGB(120, 120, 120)
                    Case corr >= -0.4
                        .Interior.Color = RGB(10, 35, 10)
                        .Font.Color = RGB(100, 200, 100)
                    Case Else
                        .Interior.Color = RGB(0, 60, 0)
                        .Font.Color = RGB(80, 220, 80)
                End Select

                If ri = ci Then
                    .Interior.Color = RGB(35, 25, 0)
                    .Font.Color = RGB(255, 192, 0)
                End If
                ' Black border
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = RGB(0, 0, 0)
                    .Weight = xlThin
                End With
            End With
        Next ci
    Next ri

    ' ¢w¢w Gold separator ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    With wsC.Range(wsC.cells(MATRIX_ROW, 1), _
                   wsC.cells(MATRIX_ROW, MATRIX_COL + nSec - 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(255, 192, 0)
        .Weight = xlThin
    End With

    ' ¢w¢w Legend ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    Dim legRow As Long: legRow = MATRIX_ROW + nSec + 2
    wsC.cells(legRow, 1).Value = "COLOUR LEGEND:"
    wsC.cells(legRow, 1).Font.Color = RGB(150, 150, 150)
    wsC.cells(legRow, 1).Font.Bold = True

    Dim legItems As Variant
    legItems = Array( _
        Array(">= 0.7  High +corr", RGB(120, 0, 0), RGB(255, 180, 180)), _
        Array("0.4~0.7 Mid  +corr", RGB(70, 15, 15), RGB(230, 130, 130)), _
        Array("-0.1~0.4 Neutral", RGB(25, 25, 25), RGB(190, 190, 190)), _
        Array("-0.4~-0.1 Mild -corr", RGB(10, 35, 10), RGB(100, 200, 100)), _
        Array("<= -0.4 High -corr", RGB(0, 60, 0), RGB(80, 220, 80)))

    Dim li As Integer
    For li = 0 To UBound(legItems)
        With wsC.cells(legRow + 1 + li, 1)
            .Value = legItems(li)(0)
            .Interior.Color = legItems(li)(1)
            .Font.Color = legItems(li)(2)
            .Font.Size = 9
            .Font.Bold = True
        End With
    Next li

    ' ¢w¢w Freeze panes ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
    wsC.cells(MATRIX_ROW + 1, MATRIX_COL).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    wsC.cells(1, 1).Select

    Application.StatusBar = "Correlation matrix done ¡X " & Format(Now, "hh:mm:ss")
    MsgBox "Done! (" & nSec & " sectors, " & nDays & " days)", vbInformation
End Sub

' ¢w¢w Fetch closing price array (nDays entries) ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
Private Function FetchPriceArray(Ticker As String, nDays As Long) As Double()
    Dim result() As Double
    ReDim result(0 To nDays - 1)

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & _
          "?interval=1d&range=" & CStr(CLng(nDays * 1.5 / 365 * 12)) & "mo"
    If CLng(nDays * 1.5) > 365 Then url = Replace(url, "range=", "range=2y&fake=")
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=1y"

    On Error Resume Next
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    On Error GoTo 0

    If http.status <> 200 Then FetchPriceArray = result: Exit Function

    Dim resp As String: resp = http.responseText
    Dim pos  As Long: pos = InStr(resp, """close"":[")
    If pos = 0 Then FetchPriceArray = result: Exit Function

    Dim arrStr As String
    arrStr = Mid(resp, pos + 9, 5000)
    arrStr = Split(arrStr, "]")(0)
    Dim parts() As String: parts = Split(arrStr, ",")

    ' Take last nDays values
    Dim total As Long: total = UBound(parts) + 1
    Dim startIdx As Long: startIdx = IIf(total > nDays, total - nDays, 0)
    Dim idx As Long: idx = 0
    Dim k As Long
    For k = startIdx To UBound(parts)
        If IsNumeric(Trim(parts(k))) And idx < nDays Then
            result(idx) = CDbl(Trim(parts(k)))
            idx = idx + 1
        End If
    Next k

    FetchPriceArray = result
End Function

' ¢w¢w Pearson correlation between two return series ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
Private Function CalcCorrelation(retData() As Double, r1 As Integer, r2 As Integer, n As Long) As Double
    Dim sumX As Double, sumY As Double, sumXY As Double
    Dim sumX2 As Double, sumY2 As Double
    Dim i As Long, cnt As Long

    For i = 0 To n - 1
        Dim x As Double: x = retData(r1, i)
        Dim y As Double: y = retData(r2, i)
        If x <> 0 Or y <> 0 Then
            sumX = sumX + x
            sumY = sumY + y
            sumXY = sumXY + x * y
            sumX2 = sumX2 + x * x
            sumY2 = sumY2 + y * y
            cnt = cnt + 1
        End If
    Next i

    If cnt < 5 Then CalcCorrelation = 0: Exit Function

    Dim num As Double: num = cnt * sumXY - sumX * sumY
    Dim den As Double: den = Sqr((cnt * sumX2 - sumX ^ 2) * (cnt * sumY2 - sumY ^ 2))
    If den = 0 Then CalcCorrelation = 0 Else CalcCorrelation = num / den
End Function

' ¢w¢w Helper: safely name a cell (optional) ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
Private Sub Name_Safe(ws As Worksheet)
    ' placeholder, not used
End Sub

' ¢w¢w Refresh with new days value ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
Sub RefreshCorrelation()
    Call BuildCorrelationMatrix
End Sub

