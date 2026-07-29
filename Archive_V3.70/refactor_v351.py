import re
import os

def rewrite_vba():
    in_file = r"C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.50.bas"
    out_file = r"C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.51.bas"
    
    with open(in_file, 'r', encoding='utf-8') as f:
        content = f.read()

    new_bp = '''Public Sub BuildPresentation()
    On Error GoTo ERR_HANDLER
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    wsMain.Unprotect "Z961814r"
    With wsMain.Range("G15:K15")
        .Merge
        .Value = ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1503) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & "..."
        .Font.Name = "Assistant"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 153)
    End With
    wsMain.Protect DrawingObjects:=False, UserInterfaceOnly:=True
    Application.ScreenUpdating = True
    DoEvents

    Dim refYear As String
    Dim yearVal As String
    Dim paramsSubtitle As String
    Dim levavName As String
    
    refYear = wsMain.Range("rngRefYear").Text
    yearVal = wsMain.Range("rngCurrentYear").Text
    paramsSubtitle = GetParamsSubtitle()
    
    levavName = ""
    Dim hi As Long
    For hi = 1 To hiddenCount
        If hiddenSheets(hi) = SHEET_AGENTS() Then levavName = hiddenExcludes(hi)
    Next hi
    
    Dim sheetList() As String
    Dim sheetCount As Long
    sheetCount = 0
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If sh.Visible = xlSheetVisible And InStr(1, "|" & SHEET_COMPANIES() & "|" & SHEET_AGENTS() & "|" & SHEET_BRANCH() & "|" & SHEET_MAINBRANCH() & "|" & SHEET_TELLERS() & "|", "|" & sh.Name & "|", vbTextCompare) > 0 Then
            sheetCount = sheetCount + 1
            ReDim Preserve sheetList(1 To sheetCount)
            sheetList(sheetCount) = sh.Name
        End If
    Next sh
    
    If sheetCount = 0 Then
        MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1492) & ChrW(1510) & ChrW(1490) & ChrW(1492), vbExclamation
        GoTo CLEAN_EXIT
    End If

    Dim ppApp As Object
    Dim ppPres As Object
    Dim ppSlide As Object
    Dim ppWeOwnApp As Boolean
    
    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Then
        Set ppApp = CreateObject("PowerPoint.Application")
        ppWeOwnApp = True
    End If
    On Error GoTo ERR_HANDLER
    If ppApp Is Nothing Then
        MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1508) & ChrW(1514) & ChrW(1497) & ChrW(1495) & ChrW(1514) & " PowerPoint", vbCritical
        GoTo CLEAN_EXIT
    End If
    
    ppApp.Visible = True
    ppApp.WindowState = 2 ' Minimized for speed and stability
    
    Set ppPres = ppApp.Presentations.Add
    ppPres.PageSetup.SlideSize = 1 ' ppSlideSizeOnScreen
    
    Dim slideW As Single
    Dim slideH As Single
    slideW = ppPres.PageSetup.SlideWidth
    slideH = ppPres.PageSetup.SlideHeight
    
    ' --- Title Slide ---
    Set ppSlide = ppPres.Slides.Add(1, 12) ' ppLayoutBlank
    Dim shpTitle As Object
    Set shpTitle = ppSlide.Shapes.AddTextbox(1, 40, slideH / 2 - 60, slideW - 80, 80)
    shpTitle.TextFrame.TextRange.Text = ChrW(1505) & ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1501) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
    shpTitle.TextFrame.TextRange.Font.Size = 44
    shpTitle.TextFrame.TextRange.Font.Bold = True
    shpTitle.TextFrame.TextRange.Font.Color.RGB = RGB(0, 51, 102)
    shpTitle.TextFrame.TextRange.ParagraphFormat.Alignment = 2 ' Center
    
    If paramsSubtitle <> "" Then
        Dim shpSub As Object
        Set shpSub = ppSlide.Shapes.AddTextbox(1, 40, slideH / 2 + 20, slideW - 80, 40)
        shpSub.TextFrame.TextRange.Text = paramsSubtitle
        shpSub.TextFrame.TextRange.Font.Size = 24
        shpSub.TextFrame.TextRange.Font.Color.RGB = RGB(100, 100, 100)
        shpSub.TextFrame.TextRange.ParagraphFormat.Alignment = 2 ' Center
    End If
    
    ' --- Total Slide ---
    BuildTotalSlide ppPres, yearVal, refYear, slideW, paramsSubtitle
    
    ' --- Comparison Slides ---
    Dim si As Long
    For si = 1 To sheetCount
        BuildCompSlides ppPres, sheetList(si), yearVal, refYear, slideW, paramsSubtitle, "", True, True
    Next si
    
    ' --- Without Levav (if exists) ---
    If levavName <> "" Then
        BuildCompSlides ppPres, SHEET_AGENTS(), yearVal, refYear, slideW, paramsSubtitle, levavName, False, False
    End If
    
    ' --- Closing Slide ---
    Set ppSlide = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12)
    Dim shpEnd As Object
    Set shpEnd = ppSlide.Shapes.AddTextbox(1, 40, slideH / 2 - 40, slideW - 80, 80)
    shpEnd.TextFrame.TextRange.Text = ChrW(1514) & ChrW(1493) & ChrW(1491) & ChrW(1492)
    shpEnd.TextFrame.TextRange.Font.Size = 48
    shpEnd.TextFrame.TextRange.Font.Bold = True
    shpEnd.TextFrame.TextRange.Font.Color.RGB = RGB(0, 51, 102)
    shpEnd.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    ' Footers
    Dim pg As Long
    For pg = 1 To ppPres.Slides.Count
        AddSlideFooter ppPres.Slides(pg), pg, ppPres.Slides.Count, slideW, slideH, GetActiveAgencyName()
    Next pg
    
    ' Save
    Dim reportsFolder As String
    Dim fsoRpt As Object
    Set fsoRpt = CreateObject("Scripting.FileSystemObject")
    reportsFolder = REPORTS_FOLDER()
    If Not fsoRpt.FolderExists(reportsFolder) Then fsoRpt.CreateFolder reportsFolder
    
    Dim presFileName As String
    presFileName = ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1492) & ChrW(1504) & ChrW(1492) & ChrW(1500) & ChrW(1492) & " " & yearVal
    Dim finalPptxPath As String
    finalPptxPath = reportsFolder & "\\" & presFileName & ".pptx"
    Dim finalPdfPath As String
    finalPdfPath = reportsFolder & "\\" & presFileName & ".pdf"
    
    On Error Resume Next
    Err.Clear
    ppPres.SaveAs finalPptxPath
    If Err.Number = 0 Then
        Err.Clear
        ppPres.SaveAs finalPdfPath, 32
    End If
    On Error GoTo ERR_HANDLER
    
    Application.ScreenUpdating = True
    
    ppApp.WindowState = 3 ' Maximized
    ppPres.Slides(ppPres.Slides.Count).Select
    AppActivate ppApp.Caption
    
    On Error Resume Next
    If ppApp.CommandBars.GetPressedMso("MinimizeRibbon") = False Then
        ppApp.CommandBars.ExecuteMso "MinimizeRibbon"
    End If
    On Error GoTo ERR_HANDLER
    
CLEAN_EXIT:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    wsMain.Unprotect "Z961814r"
    wsMain.Range("G15:K15").UnMerge
    With wsMain.Range("G15:K15")
        .Value = ""
        .Interior.Color = RGB(220, 240, 220)
    End With
    wsMain.Protect DrawingObjects:=False, UserInterfaceOnly:=True
    On Error GoTo 0
    Exit Sub
    
ERR_HANDLER:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Dim errDesc As String
    errDesc = Err.Description
    On Error Resume Next
    wsMain.Unprotect "Z961814r"
    wsMain.Range("G15:K15").UnMerge
    With wsMain.Range("G15:K15")
        .Value = ""
        .Interior.Color = RGB(220, 240, 220)
    End With
    wsMain.Protect DrawingObjects:=False, UserInterfaceOnly:=True
    If Not ppPres Is Nothing Then ppPres.Close
    If Not ppApp Is Nothing And ppWeOwnApp Then ppApp.Quit
    On Error GoTo 0
    MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & vbCrLf & errDesc, vbCritical
End Sub'''

    new_bts = '''Private Sub BuildTotalSlide(ByVal ppPres As Object, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, Optional ByVal paramsSubtitle As String = "")
    On Error GoTo ERR_HANDLER
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_MONTHS())
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim sumPR As Double, sumPC As Double, sumCR As Double, sumCC As Double
    sumPR = CDbl(ws.Cells(lastRow, 2).Value2)
    sumPC = CDbl(ws.Cells(lastRow, 3).Value2)
    sumCR = CDbl(ws.Cells(lastRow, 14).Value2)
    sumCC = CDbl(ws.Cells(lastRow, 15).Value2)

    Dim tmpWs As Worksheet
    Dim co As Object
    Dim xlCht As Object
    Application.ScreenUpdating = False
    Set tmpWs = ThisWorkbook.Worksheets.Add

    tmpWs.Cells(1, 1).Value = ""
    tmpWs.Cells(1, 2).Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & refYear
    tmpWs.Cells(1, 3).Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & yearVal
    tmpWs.Cells(1, 4).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & refYear
    tmpWs.Cells(1, 5).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & yearVal
    tmpWs.Cells(2, 1).Value = ChrW(1505) & ChrW(1499) & ChrW(1493) & ChrW(1501) & " " & ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1500)
    tmpWs.Cells(2, 2).Value = sumPR
    tmpWs.Cells(2, 3).Value = sumPC
    tmpWs.Cells(2, 4).Value = sumCR
    tmpWs.Cells(2, 5).Value = sumCC

    Set co = tmpWs.ChartObjects.Add(10, 10, 600, 400)
    Set xlCht = co.Chart
    xlCht.ChartType = 51
    xlCht.SetSourceData tmpWs.Range("A1:E2"), 2
    xlCht.HasTitle = False
    xlCht.HasLegend = True

    xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
    xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
    xlCht.SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
    xlCht.SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)

    Dim sTot As Long
    For sTot = 1 To 4
        xlCht.SeriesCollection(sTot).HasDataLabels = True
        xlCht.SeriesCollection(sTot).DataLabels.NumberFormat = "#,##0,""K"""
        xlCht.SeriesCollection(sTot).DataLabels.Font.Size = 10
        xlCht.SeriesCollection(sTot).DataLabels.Orientation = 90
    Next sTot
    xlCht.Axes(2).TickLabels.NumberFormat = "#,##0,""K"""

    Dim bCopied As Boolean
    Dim attempts As Integer
    bCopied = False
    For attempts = 1 To 5
        On Error Resume Next
        Err.Clear
        xlCht.CopyPicture 1, -4147
        If Err.Number = 0 Then
            bCopied = True
            Exit For
        End If
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        On Error GoTo ERR_HANDLER
    Next attempts
    If Not bCopied Then Err.Raise vbObjectError + 1, "BuildTotalSlide", "Failed to copy chart to clipboard"

    Application.DisplayAlerts = False
    tmpWs.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Dim ppSlide As Object
    Dim shp As Object
    Set ppSlide = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12)

    Set shp = ppSlide.Shapes.AddTextbox(1, 20, 10, slideW - 40, 50)
    shp.TextFrame.TextRange.Text = ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & ChrW(1493) & ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
    shp.TextFrame.TextRange.Font.Size = 24
    shp.TextFrame.TextRange.Font.Bold = True
    shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
    shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    shp.TextFrame.WordWrap = True

    If paramsSubtitle <> "" Then
        Set shp = ppSlide.Shapes.AddTextbox(1, 40, 52, slideW - 80, 22)
        shp.TextFrame.TextRange.Text = paramsSubtitle
        shp.TextFrame.TextRange.Font.Size = 12
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    End If

    Dim ppShape As Object
    On Error Resume Next
    Set ppShape = ppSlide.Shapes.Paste(1)
    On Error GoTo ERR_HANDLER
    If Not ppShape Is Nothing Then
        ppShape.Left = 80
        ppShape.Top = 78
        ppShape.Width = slideW - 160
        ppShape.Height = 432
    End If
    Exit Sub
ERR_HANDLER:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not tmpWs Is Nothing Then tmpWs.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "BuildTotalSlide:" & Erl, Err.Description
End Sub'''

    new_bcs = '''Private Sub BuildCompSlides(ByVal ppPres As Object, ByVal sheetName As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, Optional ByVal paramsSubtitle As String = "", Optional ByVal excludeName As String = "", Optional ByVal incDocs As Boolean = True, Optional ByVal incInsured As Boolean = True)
    On Error GoTo ERR_HANDLER
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim dataRows As Long
    dataRows = lastRow - 3

    Dim arrNames() As String, arrPremR() As Double, arrPremC() As Double, arrCommR() As Double, arrCommC() As Double
    Dim arrDocsR() As Double, arrDocsC() As Double, arrInsR() As Double, arrInsC() As Double
    Dim nItems As Long
    nItems = dataRows - 1
    If nItems < 1 Then Exit Sub

    ReDim arrNames(1 To nItems), arrPremR(1 To nItems), arrPremC(1 To nItems)
    ReDim arrCommR(1 To nItems), arrCommC(1 To nItems)
    ReDim arrDocsR(1 To nItems), arrDocsC(1 To nItems)
    ReDim arrInsR(1 To nItems), arrInsC(1 To nItems)

    Dim idx As Long, r As Long, tmpName As String
    idx = 0
    For r = 4 To lastRow - 1
        tmpName = ShortenCompanyName(Trim$(CStr(ws.Cells(r, 1).Value2)))
        If Len(excludeName) > 0 Then
            If InStr(1, tmpName, excludeName, vbTextCompare) > 0 Then GoTo NEXT_ROW_ECC
        End If
        idx = idx + 1
        If idx > nItems Then Exit For
        arrNames(idx) = tmpName
        arrPremR(idx) = CDbl(ws.Cells(r, 2).Value2)
        arrPremC(idx) = CDbl(ws.Cells(r, 3).Value2)
        arrDocsR(idx) = CDbl(ws.Cells(r, 5).Value2)
        arrDocsC(idx) = CDbl(ws.Cells(r, 6).Value2)
        arrInsR(idx) = CDbl(ws.Cells(r, 8).Value2)
        arrInsC(idx) = CDbl(ws.Cells(r, 9).Value2)
        arrCommR(idx) = CDbl(ws.Cells(r, 14).Value2)
        arrCommC(idx) = CDbl(ws.Cells(r, 15).Value2)
NEXT_ROW_ECC:
    Next r
    nItems = idx

    Dim chartItems As Long
    chartItems = nItems
    If chartItems = 0 Then Exit Sub
    If chartItems > 15 Then chartItems = 15

    Dim tmpWs As Worksheet
    Dim co As Object, xlCht As Object
    Dim ci As Long
    Application.ScreenUpdating = False
    Set tmpWs = ThisWorkbook.Worksheets.Add

    ' Helper inline logic for each chart
    Dim tPrem As String, tComm As String, tDocs As String, tIns As String
    Dim suffix As String
    suffix = ""
    If excludeName <> "" Then suffix = " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & excludeName
    tPrem = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & suffix
    tComm = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & suffix
    tDocs = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501) & suffix
    tIns = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501) & suffix

    ' --- Chart 1: Premiums ---
    tmpWs.Cells.Clear
    tmpWs.Cells(1, 1).Value = ""
    tmpWs.Cells(1, 2).Value = refYear
    tmpWs.Cells(1, 3).Value = yearVal
    For ci = 1 To chartItems
        tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
        tmpWs.Cells(ci + 1, 2).Value = arrPremR(ci)
        tmpWs.Cells(ci + 1, 3).Value = arrPremC(ci)
    Next ci
    Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
    Set xlCht = co.Chart
    xlCht.ChartType = 51
    xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3)), 2
    xlCht.HasTitle = False
    xlCht.HasLegend = True
    xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
    xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
    Dim sP As Long
    For sP = 1 To 2
        xlCht.SeriesCollection(sP).HasDataLabels = True
        xlCht.SeriesCollection(sP).DataLabels.NumberFormat = "#,##0,""K"""
        xlCht.SeriesCollection(sP).DataLabels.Font.Size = 9
        xlCht.SeriesCollection(sP).DataLabels.Orientation = 90
    Next sP
    xlCht.Axes(2).TickLabels.NumberFormat = "#,##0,""K"""
    PasteToSlide ppPres, xlCht, tPrem, yearVal, refYear, slideW, paramsSubtitle

    ' --- Chart 2: Commissions ---
    tmpWs.ChartObjects.Delete
    tmpWs.Cells.Clear
    tmpWs.Cells(1, 1).Value = ""
    tmpWs.Cells(1, 2).Value = refYear
    tmpWs.Cells(1, 3).Value = yearVal
    For ci = 1 To chartItems
        tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
        tmpWs.Cells(ci + 1, 2).Value = arrCommR(ci)
        tmpWs.Cells(ci + 1, 3).Value = arrCommC(ci)
    Next ci
    Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
    Set xlCht = co.Chart
    xlCht.ChartType = 51
    xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3)), 2
    xlCht.HasTitle = False
    xlCht.HasLegend = True
    xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
    xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
    Dim sC As Long
    For sC = 1 To 2
        xlCht.SeriesCollection(sC).HasDataLabels = True
        xlCht.SeriesCollection(sC).DataLabels.NumberFormat = "#,##0,""K"""
        xlCht.SeriesCollection(sC).DataLabels.Font.Size = 9
        xlCht.SeriesCollection(sC).DataLabels.Orientation = 90
    Next sC
    xlCht.Axes(2).TickLabels.NumberFormat = "#,##0,""K"""
    PasteToSlide ppPres, xlCht, tComm, yearVal, refYear, slideW, paramsSubtitle

    ' --- Chart 3: Documents ---
    If incDocs Then
        tmpWs.ChartObjects.Delete
        tmpWs.Cells.Clear
        tmpWs.Cells(1, 1).Value = ""
        tmpWs.Cells(1, 2).Value = refYear
        tmpWs.Cells(1, 3).Value = yearVal
        For ci = 1 To chartItems
            tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
            tmpWs.Cells(ci + 1, 2).Value = arrDocsR(ci)
            tmpWs.Cells(ci + 1, 3).Value = arrDocsC(ci)
        Next ci
        Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
        Set xlCht = co.Chart
        xlCht.ChartType = 51
        xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3)), 2
        xlCht.HasTitle = False
        xlCht.HasLegend = True
        xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(180, 130, 70)
        xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(91, 155, 213)
        Dim sD As Long
        For sD = 1 To 2
            xlCht.SeriesCollection(sD).HasDataLabels = True
            xlCht.SeriesCollection(sD).DataLabels.NumberFormat = "#,##0"
            xlCht.SeriesCollection(sD).DataLabels.Font.Size = 9
            xlCht.SeriesCollection(sD).DataLabels.Orientation = 90
        Next sD
        xlCht.Axes(2).TickLabels.NumberFormat = "#,##0"
        PasteToSlide ppPres, xlCht, tDocs, yearVal, refYear, slideW, paramsSubtitle
    End If

    ' --- Chart 4: Insured ---
    If incInsured Then
        tmpWs.ChartObjects.Delete
        tmpWs.Cells.Clear
        tmpWs.Cells(1, 1).Value = ""
        tmpWs.Cells(1, 2).Value = refYear
        tmpWs.Cells(1, 3).Value = yearVal
        For ci = 1 To chartItems
            tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
            tmpWs.Cells(ci + 1, 2).Value = arrInsR(ci)
            tmpWs.Cells(ci + 1, 3).Value = arrInsC(ci)
        Next ci
        Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
        Set xlCht = co.Chart
        xlCht.ChartType = 51
        xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3)), 2
        xlCht.HasTitle = False
        xlCht.HasLegend = True
        xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(128, 0, 128)
        xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        Dim sI2 As Long
        For sI2 = 1 To 2
            xlCht.SeriesCollection(sI2).HasDataLabels = True
            xlCht.SeriesCollection(sI2).DataLabels.NumberFormat = "#,##0"
            xlCht.SeriesCollection(sI2).DataLabels.Font.Size = 9
            xlCht.SeriesCollection(sI2).DataLabels.Orientation = 90
        Next sI2
        xlCht.Axes(2).TickLabels.NumberFormat = "#,##0"
        PasteToSlide ppPres, xlCht, tIns, yearVal, refYear, slideW, paramsSubtitle
    End If

    Application.DisplayAlerts = False
    tmpWs.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ERR_HANDLER:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not tmpWs Is Nothing Then tmpWs.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "BuildCompSlides:" & Erl, Err.Description
End Sub'''

    new_paste = '''Private Sub PasteToSlide(ByVal ppPres As Object, ByVal xlCht As Object, ByVal chartTitle As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, ByVal paramsSubtitle As String)
    Dim bCopied As Boolean
    Dim attempts As Integer
    bCopied = False
    For attempts = 1 To 5
        On Error Resume Next
        Err.Clear
        xlCht.CopyPicture 1, -4147 ' xlScreen, xlPicture
        If Err.Number = 0 Then
            bCopied = True
            Exit For
        End If
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        On Error GoTo 0
    Next attempts
    If Not bCopied Then Err.Raise vbObjectError + 1, "PasteToSlide", "Failed to copy chart to clipboard"
    
    Dim ppSlide As Object
    Set ppSlide = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12)
    Dim shp As Object
    Set shp = ppSlide.Shapes.AddTextbox(1, 20, 8, slideW - 40, 36)
    shp.TextFrame.TextRange.Text = chartTitle & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
    shp.TextFrame.TextRange.Font.Size = 20
    shp.TextFrame.TextRange.Font.Bold = True
    shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
    shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    shp.TextFrame.WordWrap = True

    If paramsSubtitle <> "" Then
        Set shp = ppSlide.Shapes.AddTextbox(1, 40, 42, slideW - 80, 22)
        shp.TextFrame.TextRange.Text = paramsSubtitle
        shp.TextFrame.TextRange.Font.Size = 12
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    End If
    
    Dim ppShape As Object
    On Error Resume Next
    Set ppShape = ppSlide.Shapes.Paste(1)
    On Error GoTo 0
    If Not ppShape Is Nothing Then
        ppShape.Left = 60
        ppShape.Top = 68
        ppShape.Width = slideW - 120
        ppShape.Height = 448
    End If
End Sub'''

    content = re.sub(r'(?sm)Public Sub BuildPresentation\(\).*?End Sub', new_bp, content)
    content = re.sub(r'(?sm)Private Sub SafeExportChart\(.*?\).*?End Sub', new_bts, content)
    content = re.sub(r'(?sm)Private Sub ExportTotalChart\(.*?\).*?End Sub', new_bcs, content)
    content = re.sub(r'(?sm)Private Sub ExportCompCharts\(.*?\).*?End Sub', new_paste, content)
    
    content = re.sub(r'(?sm)Private Sub BuildTotalSlideFromImage\(.*?\).*?End Sub', '', content)
    content = re.sub(r'(?sm)Private Sub BuildChartSlide\(.*?\).*?End Sub', '', content)

    content = content.replace('newZoom = CLng(ActiveWindow.Zoom * 1.3)', 'newZoom = CLng(ActiveWindow.Zoom * 1.2)')

    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Successfully wrote V3.51")

if __name__ == '__main__':
    rewrite_vba()
