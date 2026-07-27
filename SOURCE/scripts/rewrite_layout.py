import re
import sys

file_path = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.0.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# I will find the boundaries of the code to replace
start_marker = "    ' ---- Clear old labels from A2:D5 (no longer used) ----"
end_marker = "    ' ---- Store Hebrew message texts in Messages section (A222+) ----"

start_idx = content.find(start_marker)
end_idx = content.find(end_marker)

if start_idx == -1 or end_idx == -1:
    print("Could not find markers!")
    sys.exit(1)

new_code = """    ' ---- Clear old labels from A2:D5 (no longer used) ----
    wsMain.Range("A2:K20").ClearContents
    wsMain.Range("A2:K20").Interior.ColorIndex = xlNone
    wsMain.Range("A2:K20").Borders.LineStyle = xlNone

    ' ---- Write period lookup lists in PeriodLists section ----
    periodBaseRow = ThisWorkbook.Names("rngSection_PeriodLists").RefersToRange.Row
    wsMgmt.Cells(periodBaseRow + 1, 1).Value = "PERIOD_TYPE"
    wsMgmt.Cells(periodBaseRow + 2, 1).Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 3, 1).Value = ChrW(1495) & ChrW(1510) & ChrW(1497) & " " & ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 4, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 5, 1).Value = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 6, 1).Value = "HALF_YEAR"
    wsMgmt.Cells(periodBaseRow + 7, 1).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504) & ChrW(1492)
    wsMgmt.Cells(periodBaseRow + 8, 1).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & ChrW(1492)
    
    On Error Resume Next
    ThisWorkbook.Names("lst_period_type").Delete
    ThisWorkbook.Names("lst_half_year").Delete
    ThisWorkbook.Names("lst_quarter").Delete
    ThisWorkbook.Names("lst_month").Delete
    Err.Clear
    On Error GoTo ERR_HANDLER
    ThisWorkbook.Names.Add "lst_period_type", wsMgmt.Range(wsMgmt.Cells(periodBaseRow + 2, 1), wsMgmt.Cells(periodBaseRow + 5, 1))
    ThisWorkbook.Names.Add "lst_half_year", wsMgmt.Range(wsMgmt.Cells(periodBaseRow + 7, 1), wsMgmt.Cells(periodBaseRow + 8, 1))
    
    wsMgmt.Cells(periodBaseRow + 10, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503)
    wsMgmt.Cells(periodBaseRow + 11, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 12, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 13, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497)
    ThisWorkbook.Names.Add "lst_quarter", wsMgmt.Range(wsMgmt.Cells(periodBaseRow + 10, 1), wsMgmt.Cells(periodBaseRow + 13, 1))
    
    wsMgmt.Cells(periodBaseRow + 15, 1).Value = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 16, 1).Value = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 17, 1).Value = ChrW(1502) & ChrW(1512) & ChrW(1509)
    wsMgmt.Cells(periodBaseRow + 18, 1).Value = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
    wsMgmt.Cells(periodBaseRow + 19, 1).Value = ChrW(1502) & ChrW(1488) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 20, 1).Value = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 21, 1).Value = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 22, 1).Value = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496)
    wsMgmt.Cells(periodBaseRow + 23, 1).Value = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 24, 1).Value = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 25, 1).Value = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 26, 1).Value = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    ThisWorkbook.Names.Add "lst_month", wsMgmt.Range(wsMgmt.Cells(periodBaseRow + 15, 1), wsMgmt.Cells(periodBaseRow + 26, 1))
    
    ' ---- Layout setup for Home Sheet ----
    ' Parameters Table (F2:G10)
    wsMain.Range("G2").Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512)
    wsMain.Range("F2").Value = ChrW(1506) & ChrW(1512) & ChrW(1498)
    Set hdrRng = wsMain.Range("F2:G2")
    hdrRng.Interior.Color = RGB(0, 100, 0)
    hdrRng.Font.Color = RGB(255, 255, 255)
    hdrRng.Font.Bold = True
    hdrRng.Font.Size = 12
    hdrRng.HorizontalAlignment = xlCenter
    
    wsMain.Range("G3").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505)
    wsMain.Range("G4").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1499) & ChrW(1495) & ChrW(1497) & ChrW(1514)
    wsMain.Range("G5").Value = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
    wsMain.Range("G6").Value = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
    wsMain.Range("G7").Value = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1514) & ChrW(1488) & ChrW(1512) & ChrW(1497) & ChrW(1498)
    wsMain.Range("G8").Value = ChrW(1495) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1498) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497)
    wsMain.Range("G9").Value = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1495) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1498)
    wsMain.Range("G10").Value = ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
    
    wsMain.Range("G3:G10").Font.Color = RGB(0, 70, 140)
    wsMain.Range("G3:G10").Font.Bold = True
    wsMain.Range("F3:G10").Interior.Color = RGB(255, 245, 230)
    wsMain.Range("F3:G10").HorizontalAlignment = xlCenter
    wsMain.Range("G3:G10").HorizontalAlignment = xlRight
    
    On Error Resume Next
    ThisWorkbook.Names("rngCurrentYear").Delete
    ThisWorkbook.Names("rngBaseYear").Delete
    ThisWorkbook.Names("rngPeriodType").Delete
    ThisWorkbook.Names("rngPeriodValue").Delete
    ThisWorkbook.Names("rngDateType").Delete
    ThisWorkbook.Names("rngFilterType").Delete
    ThisWorkbook.Names("rngFilterValue").Delete
    ThisWorkbook.Names("rngClientName").Delete
    Err.Clear
    On Error GoTo ERR_HANDLER
    
    ThisWorkbook.Names.Add "rngBaseYear", wsMain.Range("F3")
    ThisWorkbook.Names.Add "rngCurrentYear", wsMain.Range("F4")
    ThisWorkbook.Names.Add "rngPeriodType", wsMain.Range("F5")
    ThisWorkbook.Names.Add "rngPeriodValue", wsMain.Range("F6")
    ThisWorkbook.Names.Add "rngDateType", wsMain.Range("F7")
    ThisWorkbook.Names.Add "rngFilterType", wsMain.Range("F8")
    ThisWorkbook.Names.Add "rngFilterValue", wsMain.Range("F9")
    ThisWorkbook.Names.Add "rngClientName", wsMain.Range("F10")
    
    ' Set defaults
    If IsEmpty(wsMain.Range("F5").Value) Then wsMain.Range("F5").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    If IsEmpty(wsMain.Range("F8").Value) Then wsMain.Range("F8").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
    If IsEmpty(wsMain.Range("F10").Value) Then wsMain.Range("F10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
    
    On Error Resume Next
    wsMain.Range("F5").Validation.Delete
    wsMain.Range("F5").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_period_type"
    wsMain.Range("F6").Validation.Delete
    wsMain.Range("F7").Validation.Delete
    dateTypeList = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493) & "," & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1514) & " " & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495)
    wsMain.Range("F7").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=dateTypeList
    wsMain.Range("F8").Validation.Delete
    filterTypeList = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497) & "," & H_COMPANY() & "," & H_TELLER() & "," & H_AGENT() & "," & H_BRANCH() & "," & H_BRANCH() & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
    wsMain.Range("F8").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=filterTypeList
    wsMain.Range("F9").Validation.Delete
    Err.Clear
    On Error GoTo ERR_HANDLER
    
    ' ---- Currency Table J2:K4 ----
    wsMain.Range("K2").Value = ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)
    wsMain.Range("J2").Value = ChrW(1513) & ChrW(1506) & ChrW(1512)
    wsMain.Range("J2:K2").Interior.Color = RGB(0, 100, 0)
    wsMain.Range("J2:K2").Font.Color = RGB(255, 255, 255)
    wsMain.Range("J2:K2").Font.Bold = True
    wsMain.Range("J2:K2").Font.Size = 12
    wsMain.Range("J2:K2").HorizontalAlignment = xlCenter
    wsMain.Range("K3").Value = "$" & " " & ChrW(1491) & ChrW(1493) & ChrW(1500) & ChrW(1512)
    wsMain.Range("K4").Value = ChrW(8364) & " " & ChrW(1488) & ChrW(1497) & ChrW(1512) & ChrW(1493)
    wsMain.Range("K3:K4").Font.Bold = True
    wsMain.Range("K3:K4").HorizontalAlignment = xlCenter
    wsMain.Range("J3:K4").Interior.Color = RGB(255, 245, 230)
    
    On Error Resume Next
    ThisWorkbook.Names("rngDOLAR").Delete
    Err.Clear
    On Error GoTo ERR_HANDLER
    ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("J3")
    
    curRate = FetchBOIRate()
    If curRate > 0 Then
        wsMain.Range("J3").Value = curRate
    ElseIf IsEmpty(wsMain.Range("J3").Value) Then
        wsMain.Range("J3").Value = 3.6
    End If
    wsMain.Range("J3").NumberFormat = "0.0000"
    
    eurRate = FetchBOIMonthlyAvg("EUR", "")
    If eurRate > 0 Then
        wsMain.Range("J4").Value = eurRate
    ElseIf IsEmpty(wsMain.Range("J4").Value) Then
        wsMain.Range("J4").Value = 3.9
    End If
    wsMain.Range("J4").NumberFormat = "0.0000"
    wsMain.Range("J3:J4").HorizontalAlignment = xlCenter
    
    matachMsg = ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1497) & ChrW(1501) & " " & _
        ChrW(1489) & ChrW(1502) & ChrW(1496) & ChrW(34) & ChrW(1495) & " " & _
        ChrW(1502) & ChrW(1514) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & _
        ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & _
        ChrW(1492) & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & _
        ChrW(1492) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & _
        ChrW(1489) & ChrW(1504) & ChrW(1511) & " " & _
        ChrW(1497) & ChrW(1513) & ChrW(1512) & ChrW(1488) & ChrW(1500) & " " & _
        ChrW(1500) & "-15 " & _
        ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & " " & _
        ChrW(1492) & ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)
    wsMain.Range("J5:K7").Merge
    wsMain.Range("J5").Value = matachMsg
    wsMain.Range("J5").WrapText = True
    wsMain.Range("J5").Font.Size = 9
    wsMain.Range("J5").HorizontalAlignment = xlCenter
    wsMain.Range("J5").VerticalAlignment = xlCenter
    wsMain.Range("J5:K7").Interior.Color = RGB(230, 245, 255)
    
    ' ---- Formatting Columns ----
    wsMain.Columns("A").ColumnWidth = 8
    wsMain.Columns("B").ColumnWidth = 8
    wsMain.Columns("C").ColumnWidth = 18
    wsMain.Columns("D").ColumnWidth = 8
    wsMain.Columns("E").ColumnWidth = 8
    wsMain.Columns("F").ColumnWidth = 14
    wsMain.Columns("G").ColumnWidth = 14
    wsMain.Columns("H").ColumnWidth = 8
    wsMain.Columns("I").ColumnWidth = 8
    wsMain.Columns("J").ColumnWidth = 12
    wsMain.Columns("K").ColumnWidth = 12
    wsMain.Rows("1:24").RowHeight = 22
    wsMain.Rows("11:13").RowHeight = 35
    
    ' ---- Borders ----
    For Each cellBorder In wsMain.Range("F2:G10")
        For Each bEdge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            cellBorder.Borders(bEdge).LineStyle = xlContinuous
            cellBorder.Borders(bEdge).Color = RGB(0, 70, 140)
            cellBorder.Borders(bEdge).Weight = xlThin
        Next bEdge
    Next cellBorder
    For Each cellBorder In wsMain.Range("J2:K4")
        For Each bEdge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            cellBorder.Borders(bEdge).LineStyle = xlContinuous
            cellBorder.Borders(bEdge).Color = RGB(0, 70, 140)
            cellBorder.Borders(bEdge).Weight = xlThin
        Next bEdge
    Next cellBorder
    For Each cellBorder In wsMain.Range("J5:K7")
        For Each bEdge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            cellBorder.Borders(bEdge).LineStyle = xlContinuous
            cellBorder.Borders(bEdge).Color = RGB(0, 70, 140)
            cellBorder.Borders(bEdge).Weight = xlThin
        Next bEdge
    Next cellBorder
    
    ' ---- Remove ALL old buttons ----
    On Error Resume Next
    For Each s In wsMain.Shapes
        s.Delete
    Next s
    Err.Clear
    On Error GoTo ERR_HANDLER
    
    ' ---- Buttons Column C (1 to 6) ----
    Dim btnLeft As Double, btnW As Double, btnH As Double
    btnW = 140
    btnH = 35
    btnLeft = wsMain.Range("C2").Left + (wsMain.Range("C2").Width - btnW) / 2
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top, btnW, btnH)
    shp.Name = "btnBuildReview"
    shp.Fill.ForeColor.RGB = blueClr
    shp.TextFrame2.TextRange.Text = "1 - " & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "BuildReview"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C4").Top, btnW, btnH)
    shp.Name = "btnApplyCorrections"
    shp.Fill.ForeColor.RGB = RGB(0, 150, 80)
    shp.TextFrame2.TextRange.Text = "2 - " & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(34) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "ApplyCorrectionsAndBuildReports"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C6").Top, btnW, btnH)
    shp.Name = "btnBuildPresentation"
    shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
    shp.TextFrame2.TextRange.Text = "3 - " & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "BuildPresentation"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C8").Top, btnW, btnH)
    shp.Name = "btnSaveReports"
    shp.Fill.ForeColor.RGB = blueClr
    shp.TextFrame2.TextRange.Text = "4 - " & ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "SaveReportsToFolder"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C10").Top, btnW, btnH)
    shp.Name = "btnViewReports"
    shp.Fill.ForeColor.RGB = RGB(0, 150, 80)
    shp.TextFrame2.TextRange.Text = "5 - " & ChrW(1510) & ChrW(1508) & ChrW(1497) & ChrW(1497) & ChrW(1492) & " " & ChrW(1489) & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "ViewReportsFolder"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C12").Top, btnW, btnH)
    shp.Name = "btnNewClients"
    shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
    shp.TextFrame2.TextRange.Text = "6 - " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NewClients"
    
    wsMain.Range("C14").Value = ChrW(1508) & ChrW(1493) & ChrW(1514) & ChrW(1495) & " " & ChrW(1506) & ChrW(34) & ChrW(1497) & " " & ChrW(1505) & ChrW(1497) & ChrW(1500) & ChrW(1489) & ChrW(1503) & " " & ChrW(1496) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1500) & ChrW(1493) & ChrW(1490) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " 054-6677396"
    wsMain.Range("C14").Font.Color = RGB(0, 120, 0)
    wsMain.Range("C14").Font.Bold = True
    wsMain.Range("C14").Font.Size = 10
    wsMain.Range("C14").HorizontalAlignment = xlCenter
    
    ' ---- Helper buttons ----
    Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G11").Left + 5, wsMain.Range("G11").Top + 2, 60, 25)
    btnSearch.Name = "btnSearchClient"
    btnSearch.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)
    btnSearch.TextFrame2.TextRange.Font.Size = 10
    btnSearch.TextFrame2.TextRange.Font.Bold = msoTrue
    btnSearch.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btnSearch.Fill.ForeColor.RGB = RGB(0, 150, 80)
    btnSearch.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btnSearch.OnAction = "SearchClientName"
    
    Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F11").Left + 5, wsMain.Range("F11").Top + 2, 60, 25)
    btnAll.Name = "btnAllClients"
    btnAll.TextFrame2.TextRange.Text = ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1501)
    btnAll.TextFrame2.TextRange.Font.Size = 10
    btnAll.TextFrame2.TextRange.Font.Bold = msoTrue
    btnAll.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btnAll.Fill.ForeColor.RGB = blueClr
    btnAll.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btnAll.OnAction = "ResetClientFilter"
    
    Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G12").Left, wsMain.Range("G12").Top + 5, wsMain.Range("G12").Width + wsMain.Range("F12").Width, 25)
    btnReset.Name = "btnResetDefaults"
    btnReset.TextFrame2.TextRange.Text = ChrW(1488) & ChrW(1508) & ChrW(1505) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
    btnReset.TextFrame2.TextRange.Font.Size = 11
    btnReset.TextFrame2.TextRange.Font.Bold = msoTrue
    btnReset.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btnReset.Fill.ForeColor.RGB = RGB(180, 50, 50)
    btnReset.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btnReset.OnAction = "ResetHomeDefaults"
    
    ' Show/Hide Sheets in J9
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("J9").Left, wsMain.Range("J9").Top, 160, 30)
    shp.Name = "btnShowHidden"
    shp.Fill.ForeColor.RGB = RGB(80, 80, 80)
    shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & "/" & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1512) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "ToggleHiddenSheets"
    
    ' ---- rngFILES_FOLDER and rngREPORTS_FOLDER ----
    On Error Resume Next
    existingPath = Trim$(CStr(ThisWorkbook.Names("rngFILES_FOLDER").RefersToRange.Value2))
    If existingPath = "" Then
        ThisWorkbook.Names("rngFILES_FOLDER").Delete
        Err.Clear
        ThisWorkbook.Names.Add "rngFILES_FOLDER", wsMgmt.Range("B176")
    End If
    existingPath = Trim$(CStr(ThisWorkbook.Names("rngREPORTS_FOLDER").RefersToRange.Value2))
    If existingPath = "" Then
        ThisWorkbook.Names("rngREPORTS_FOLDER").Delete
        Err.Clear
        ThisWorkbook.Names.Add "rngREPORTS_FOLDER", wsMgmt.Range("B177")
    End If
    Err.Clear
    On Error GoTo ERR_HANDLER
    
"""

final_content = content[:start_idx] + new_code + content[end_idx:]

with open(file_path, "w", encoding="utf-8") as f:
    f.write(final_content)

print("Patch applied successfully.")
