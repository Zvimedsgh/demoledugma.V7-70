Attribute VB_Name = "modCharts"
Option Explicit

Public Sub BuildAllChartSheets()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    BuildChartSheet SheetCompanies(), "charts_companies", CatCompanies(), True
    BuildChartSheet SheetBranch(), "charts_branches", CatBranches(), False
    BuildChartSheet SheetMainBranch(), "charts_main_branch", CatMainBranch(), False
    BuildChartSheet SheetTellers(), "charts_tellers", CatTellers(), False
    BuildChartSheet SheetAgents(), "charts_agents", CatAgents(), False

    If SheetExists(SheetAgentsNoLevav()) Then
        BuildChartSheet SheetAgentsNoLevav(), "charts_agents_no_levav", CatAgentsNoLevav(), False
    End If

    If SheetExists(SheetDrachim()) Then
        BuildChartSheet SheetDrachim(), "charts_drachim", CatDrachim(), False
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Chart sheets were created successfully", vbInformation

End Sub

Private Sub BuildChartSheet(ByVal sourceSheetName As String, ByVal chartSheetName As String, ByVal categoryTitle As String, ByVal cleanCompanyNames As Boolean)

    Dim wsSrc As Worksheet
    Dim wsChart As Worksheet
    Dim lastRow As Long
    Dim baseYear As String
    Dim currYear As String

    If Not SheetExists(sourceSheetName) Then Exit Sub

    baseYear = Trim$(CStr(Range("rngBaseYear").Value))
    currYear = Trim$(CStr(Range("rngCurrentYear").Value))

    Set wsSrc = ThisWorkbook.Worksheets(sourceSheetName)

    DeleteSheetIfExists chartSheetName
    Set wsChart = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsChart.Name = chartSheetName

    lastRow = CopyChartData(wsSrc, wsChart, cleanCompanyNames)

    If lastRow < 2 Then Exit Sub

    CreateChartOnSheet wsChart, _
        TitlePremium() & " " & TitleBy() & " " & categoryTitle & " - " & baseYear & " " & TitleVs() & " " & currYear, _
        2, 3, baseYear, currYear, 20, 30

    CreateChartOnSheet wsChart, _
        TitleCommission() & " " & TitleBy() & " " & categoryTitle & " - " & baseYear & " " & TitleVs() & " " & currYear, _
        4, 5, baseYear, currYear, 20, 430

    wsChart.DisplayRightToLeft = True
    wsChart.Columns.AutoFit

End Sub

Private Function CopyChartData(ByVal wsSrc As Worksheet, ByVal wsOut As Worksheet, ByVal cleanCompanyNames As Boolean) As Long

    Dim lastSrcRow As Long
    Dim r As Long
    Dim outRow As Long
    Dim labelText As String

    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    wsOut.Cells(1, 1).Value = "Label"
    wsOut.Cells(1, 2).Value = "Premium_Base"
    wsOut.Cells(1, 3).Value = "Premium_Current"
    wsOut.Cells(1, 4).Value = "Commission_Base"
    wsOut.Cells(1, 5).Value = "Commission_Current"

    outRow = 2

    For r = 2 To lastSrcRow

        labelText = Trim$(CStr(wsSrc.Cells(r, 1).Value))

        If labelText <> "" Then
            If Not IsTotalRow(labelText) Then

                If cleanCompanyNames Then
                    labelText = CleanCompanyName(labelText)
                End If

                wsOut.Cells(outRow, 1).Value = labelText
                wsOut.Cells(outRow, 2).Value = wsSrc.Cells(r, 2).Value
                wsOut.Cells(outRow, 3).Value = wsSrc.Cells(r, 3).Value
                wsOut.Cells(outRow, 4).Value = wsSrc.Cells(r, 14).Value
                wsOut.Cells(outRow, 5).Value = wsSrc.Cells(r, 15).Value

                outRow = outRow + 1

            End If
        End If

    Next r

    CopyChartData = outRow - 1

End Function

Private Sub CreateChartOnSheet(ByVal ws As Worksheet, ByVal chartTitle As String, ByVal baseCol As Long, ByVal currCol As Long, _
                               ByVal baseYear As String, ByVal currYear As String, ByVal leftPos As Double, ByVal topPos As Double)

    Dim lastRow As Long
    Dim ch As ChartObject

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set ch = ws.ChartObjects.Add(leftPos, topPos, 900, 360)

    With ch.Chart

        .ChartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.Text = chartTitle

        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = baseYear
        .SeriesCollection(1).Values = ws.Range(ws.Cells(2, baseCol), ws.Cells(lastRow, baseCol))
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))

        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = currYear
        .SeriesCollection(2).Values = ws.Range(ws.Cells(2, currCol), ws.Cells(lastRow, currCol))
        .SeriesCollection(2).XValues = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        .Axes(xlCategory).TickLabels.Orientation = 45
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"

    End With

End Sub

Private Sub DeleteSheetIfExists(ByVal sheetName As String)

    If SheetExists(sheetName) Then
        ThisWorkbook.Worksheets(sheetName).Delete
    End If

End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function

Private Function IsTotalRow(ByVal txt As String) As Boolean
    IsTotalRow = (InStr(1, txt, H("1505 1492"), vbTextCompare) > 0)
End Function

Private Function CleanCompanyName(ByVal txt As String) As String

    txt = Replace(txt, H("1495 1489 1512 1492 32 1500 1489 1497 1496 1493 1495 32 1489 1506 34 1502"), "")
    txt = Replace(txt, H("1495 1489 1512 1492 32 1500 1489 1497 1496 1493 1495"), "")
    txt = Replace(txt, H("1489 1506 34 1502"), "")
    txt = Replace(txt, H("1495 1489 1512 1514"), "")

    CleanCompanyName = Trim$(txt)

End Function

Private Function H(ByVal codes As String) As String

    Dim arr() As String
    Dim i As Long
    Dim s As String

    arr = Split(codes, " ")

    For i = LBound(arr) To UBound(arr)
        s = s & ChrW(CLng(arr(i)))
    Next i

    H = s

End Function

Private Function SheetCompanies() As String
    SheetCompanies = H("1495 1489 1512 1493 1514")
End Function

Private Function SheetBranch() As String
    SheetBranch = H("1506 1504 1508 1497 1501")
End Function

Private Function SheetMainBranch() As String
    SheetMainBranch = H("1506 1504 1507 32 1502 1512 1499 1494")
End Function

Private Function SheetTellers() As String
    SheetTellers = H("1496 1500 1512 1497 1501")
End Function

Private Function SheetAgents() As String
    SheetAgents = H("1505 1493 1499 1504 1497 1501")
End Function

Private Function SheetAgentsNoLevav() As String
    SheetAgentsNoLevav = H("1505 1493 1499 1504 1497 1501 32 1500 1500 1488 32 1500 1489 1489")
End Function

Private Function SheetDrachim() As String
    SheetDrachim = H("1491 1512 1499 1497 1501")
End Function

Private Function CatCompanies() As String
    CatCompanies = H("1495 1489 1512 1493 1514")
End Function

Private Function CatBranches() As String
    CatBranches = H("1506 1504 1508 1497 1501")
End Function

Private Function CatMainBranch() As String
    CatMainBranch = H("1506 1504 1507 32 1502 1512 1499 1494")
End Function

Private Function CatTellers() As String
    CatTellers = H("1496 1500 1512 1497 1501")
End Function

Private Function CatAgents() As String
    CatAgents = H("1505 1493 1499 1504 1497 1501")
End Function

Private Function CatAgentsNoLevav() As String
    CatAgentsNoLevav = H("1505 1493 1499 1504 1497 1501 32 1500 1500 1488 32 1500 1489 1489")
End Function

Private Function CatDrachim() As String
    CatDrachim = H("1491 1512 1499 1497 1501")
End Function

Private Function TitlePremium() As String
    TitlePremium = H("1508 1512 1502 1497 1493 1514")
End Function

Private Function TitleCommission() As String
    TitleCommission = H("1506 1502 1500 1493 1514")
End Function

Private Function TitleBy() As String
    TitleBy = H("1500 1508 1497")
End Function

Private Function TitleVs() As String
    TitleVs = H("1502 1493 1500")
End Function

