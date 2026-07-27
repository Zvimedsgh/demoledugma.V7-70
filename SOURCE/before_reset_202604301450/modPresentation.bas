Attribute VB_Name = "modPresentation"
Option Explicit

Public Sub BuildLevavPresentationFull_OLD()

    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptPath As String
    Dim pdfPath As String
    Dim baseYear As String
    Dim currYear As String

    baseYear = Trim$(CStr(Range("rngBaseYear").Value))
    currYear = Trim$(CStr(Range("rngCurrentYear").Value))

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Add

    AddTitleSlide pptPres, baseYear, currYear

    AddSheetCharts pptPres, SheetMonths(), "Months", baseYear, currYear
    AddSheetCharts pptPres, SheetCompanies(), "Companies", baseYear, currYear
    AddSheetCharts pptPres, SheetMainBranch(), "Main branch", baseYear, currYear
    AddSheetCharts pptPres, SheetTellers(), "Tellers", baseYear, currYear
    AddSheetCharts pptPres, SheetAgents(), "Agents", baseYear, currYear

    If SheetExists(SheetAgentsNoLevav()) Then
        AddSheetCharts pptPres, SheetAgentsNoLevav(), "Agents without Levav", baseYear, currYear
    End If

    If SheetExists(SheetDrachim()) Then
        AddSheetCharts pptPres, SheetDrachim(), "Drachim", baseYear, currYear
    End If

    pptPath = ThisWorkbook.Path & "\Levav_Presentation_" & Format(Now, "yyyymmdd_hhmm") & ".pptx"
    pdfPath = ThisWorkbook.Path & "\Levav_Presentation_" & Format(Now, "yyyymmdd_hhmm") & ".pdf"

    pptPres.SaveAs pptPath
    pptPres.SaveAs pdfPath, 32

    MsgBox "Presentation and PDF saved:" & vbCrLf & pptPath & vbCrLf & pdfPath, vbInformation

End Sub

Private Sub AddTitleSlide(ByVal pptPres As Object, ByVal baseYear As String, ByVal currYear As String)

    Dim sld As Object
    Set sld = pptPres.Slides.Add(pptPres.Slides.Count + 1, 1)

    sld.Shapes(1).TextFrame.TextRange.Text = "Levav management presentation"
    sld.Shapes(2).TextFrame.TextRange.Text = baseYear & " vs " & currYear

End Sub

Private Sub AddSheetCharts(ByVal pptPres As Object, ByVal sheetName As String, ByVal titleName As String, ByVal baseYear As String, ByVal currYear As String)

    Dim ws As Worksheet
    Dim chPrem As ChartObject
    Dim chComm As ChartObject

    If Not SheetExists(sheetName) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(sheetName)

    DeleteCharts ws

    Set chPrem = CreateMetricChart(ws, titleName & " - Premium", 2, 3, baseYear, currYear, 100, 40)
    Set chComm = CreateMetricChart(ws, titleName & " - Commission", 14, 15, baseYear, currYear, 100, 430)

    AddChartSlide pptPres, chPrem, titleName & " - Premium"
    AddChartSlide pptPres, chComm, titleName & " - Commission"

End Sub

Private Function CreateMetricChart(ByVal ws As Worksheet, ByVal chartTitle As String, ByVal colBase As Long, ByVal colCurr As Long, _
                                   ByVal baseYear As String, ByVal currYear As String, ByVal leftPos As Double, ByVal topPos As Double) As ChartObject

    Dim lastRow As Long
    Dim helperCol As Long
    Dim r As Long
    Dim ch As ChartObject

    lastRow = LastDataRowNoTotal(ws)
    helperCol = 30

    For r = 2 To lastRow
        If ws.Name = SheetCompanies() Then
            ws.Cells(r, helperCol).Value = CleanCompanyName(CStr(ws.Cells(r, 1).Value))
        Else
            ws.Cells(r, helperCol).Value = ws.Cells(r, 1).Value
        End If
    Next r

    Set ch = ws.ChartObjects.Add(leftPos, topPos, 720, 340)

    With ch.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.Text = chartTitle

        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = baseYear
        .SeriesCollection(1).Values = ws.Range(ws.Cells(2, colBase), ws.Cells(lastRow, colBase))
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(2, helperCol), ws.Cells(lastRow, helperCol))

        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = currYear
        .SeriesCollection(2).Values = ws.Range(ws.Cells(2, colCurr), ws.Cells(lastRow, colCurr))
        .SeriesCollection(2).XValues = ws.Range(ws.Cells(2, helperCol), ws.Cells(lastRow, helperCol))

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlCategory).TickLabels.Orientation = 45
    End With

    Set CreateMetricChart = ch

End Function

Private Sub AddChartSlide(ByVal pptPres As Object, ByVal ch As ChartObject, ByVal slideTitle As String)

    Dim sld As Object
    Dim shp As Object

    Set sld = pptPres.Slides.Add(pptPres.Slides.Count + 1, 11)

    sld.Shapes(1).TextFrame.TextRange.Text = slideTitle

    ch.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    sld.Shapes.Paste
    Set shp = sld.Shapes(sld.Shapes.Count)

    With shp
        .LockAspectRatio = -1
        .Width = 620
        .Left = 55
        .Top = 95
    End With

End Sub

Private Sub DeleteCharts(ByVal ws As Worksheet)

    Dim ch As ChartObject
    For Each ch In ws.ChartObjects
        ch.Delete
    Next ch

End Sub

Private Function LastDataRowNoTotal(ByVal ws As Worksheet) As Long

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Do While r >= 2
        If IsTotalRow(CStr(ws.Cells(r, 1).Value)) Then
            r = r - 1
        Else
            Exit Do
        End If
    Loop

    LastDataRowNoTotal = r

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

Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function

Private Function SheetCompanies() As String
    SheetCompanies = H("1495 1489 1512 1493 1514")
End Function

Private Function SheetMonths() As String
    SheetMonths = H("1495 1493 1491 1513 1497 1501")
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

