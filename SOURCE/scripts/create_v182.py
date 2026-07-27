import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.181.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.182.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update the loop near 5905
old_loop_1 = """For Each ws In ThisWorkbook.Worksheets
If UCase$(ws.Name) <> UCase$(ctrlName) Then
ws.Visible = xlSheetVeryHidden
End If
Next ws"""

new_loop_1 = """For Each ws In ThisWorkbook.Worksheets
If UCase$(ws.Name) <> UCase$(ctrlName) Then
    Dim hideIt As Boolean
    hideIt = True
    
    If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
        If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then hideIt = False
    End If
    If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
        If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA2").Value) <> "YES" Then hideIt = False
    End If
    
    If hideIt Then ws.Visible = xlSheetVeryHidden
End If
Next ws"""

# We need to make sure we replace exactly where it appears.
# Because the source has line numbers sometimes or indentation, let's use regex or exact match
import re

content = re.sub(r'For Each ws In ThisWorkbook\.Worksheets\s*\n\s*If UCase\$\(ws\.Name\) <> UCase\$\(ctrlName\) Then\s*\n\s*ws\.Visible = xlSheetVeryHidden\s*\n\s*End If\s*\n\s*Next ws', new_loop_1, content)


# 2. Update the loop in CheckUserPermissions (around 7594)
old_loop_2 = """For Each ws In ThisWorkbook.Worksheets
If ws.Name <> homeSheetName And ws.Name <> ncSheetName Then
ws.Visible = xlSheetVeryHidden
End If
Next ws"""

new_loop_2 = """For Each ws In ThisWorkbook.Worksheets
If ws.Name <> homeSheetName And ws.Name <> ncSheetName Then
    Dim hideIt2 As Boolean
    hideIt2 = True
    
    If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
        If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA1").Value) <> "YES" Then hideIt2 = False
    End If
    If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
        If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then hideIt2 = False
    End If
    
    If hideIt2 Then ws.Visible = xlSheetVeryHidden
End If
Next ws"""

content = re.sub(r'For Each ws In ThisWorkbook\.Worksheets\s*\n\s*If ws\.Name <> homeSheetName And ws\.Name <> ncSheetName Then\s*\n\s*ws\.Visible = xlSheetVeryHidden\s*\n\s*End If\s*\n\s*Next ws', new_loop_2, content)

# 3. Add the macros at the end, right before the LogDebug method
macros = """' -------------------------------------------------------------------------
' INSTRUCTION SHEETS
' -------------------------------------------------------------------------
Public Sub CreateInstructionSheets()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    Dim wsInstall As Worksheet
    Dim wsOp As Worksheet
    Dim sInstall As String
    Dim sOp As String
    
    sInstall = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
    sOp = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
    
    ' Create or get Install Sheet
    Set wsInstall = Nothing
    Set wsInstall = ThisWorkbook.Worksheets(sInstall)
    If wsInstall Is Nothing Then
        Set wsInstall = ThisWorkbook.Worksheets.Add(After:=wsMain)
        wsInstall.Name = sInstall
    End If
    wsInstall.DisplayRightToLeft = True
    wsInstall.Cells.Clear
    
    ' Create or get Op Sheet
    Set wsOp = Nothing
    Set wsOp = ThisWorkbook.Worksheets(sOp)
    If wsOp Is Nothing Then
        Set wsOp = ThisWorkbook.Worksheets.Add(After:=wsInstall)
        wsOp.Name = sOp
    End If
    wsOp.DisplayRightToLeft = True
    wsOp.Cells.Clear
    
    ' --- Setup Install Sheet ---
    With wsInstall
        .Range("B2").Value = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1502) & ChrW(1492) & ChrW(1497) & ChrW(1512) & ChrW(1493) & ChrW(1514)
        .Range("B2").Font.Size = 24
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Color = RGB(0, 51, 102)
        
        .Range("B4").Value = ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1499) & ChrW(1497) & ChrW(1501) & " " & ChrW(1492) & ChrW(1489) & ChrW(1488) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489) & "! " & ChrW(1499) & ChrW(1491) & ChrW(1497) & " " & ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1500) & ChrW(1506) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1489) & ChrW(1510) & ChrW(1493) & ChrW(1512) & ChrW(1492) & " " & ChrW(1514) & ChrW(1511) & ChrW(1497) & ChrW(1504) & ChrW(1492) & ", " & ChrW(1488) & ChrW(1504) & ChrW(1488) & " " & ChrW(1489) & ChrW(1510) & ChrW(1506) & ChrW(1493) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1513) & ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1501) & " " & ChrW(1492) & ChrW(1489) & ChrW(1488) & ChrW(1497) & ChrW(1501) & ":"
        .Range("B4").Font.Size = 14
        .Range("B4").Font.Bold = True
        
        .Range("B6").Value = "1. " & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1493) & ChrW(1509) & " " & ChrW(1492) & ChrW(1511) & ChrW(1489) & ChrW(1510) & ChrW(1497) & ChrW(1501) & ":"
        .Range("C6").Value = ChrW(1493) & ChrW(1491) & ChrW(1488) & ChrW(1493) & " " & ChrW(1513) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1510) & ChrW(1514) & ChrW(1501) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1492) & ChrW(1511) & ChrW(1489) & ChrW(1510) & ChrW(1497) & ChrW(1501) & " " & ChrW(1502) & ChrW(1514) & ChrW(1493) & ChrW(1498) & " " & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1510) & " " & ChrW(1492) & "-ZIP " & ChrW(1500) & ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1492) & " " & ChrW(1511) & ChrW(1489) & ChrW(1493) & ChrW(1506) & ChrW(1492) & " " & ChrW(1489) & ChrW(1502) & ChrW(1495) & ChrW(1513) & ChrW(1489) & " (" & ChrW(1493) & ChrW(1500) & ChrW(1488) & " " & ChrW(1508) & ChrW(1514) & ChrW(1495) & ChrW(1514) & ChrW(1501) & " " & ChrW(1497) & ChrW(1513) & ChrW(1497) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1514) & ChrW(1493) & ChrW(1498) & " " & ChrW(1492) & "-ZIP)."
        .Range("B7").Value = "2. " & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500) & " " & ChrW(1495) & ChrW(1505) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492) & " (Unblock):"
        .Range("C7").Value = ChrW(1500) & ChrW(1495) & ChrW(1510) & ChrW(1493) & " " & ChrW(1511) & ChrW(1500) & ChrW(1497) & ChrW(1511) & " " & ChrW(1497) & ChrW(1502) & ChrW(1504) & ChrW(1497) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1510) & " " & ChrW(1492) & ChrW(1488) & ChrW(1511) & ChrW(1505) & ChrW(1500) & " -> '" & ChrW(1502) & ChrW(1488) & ChrW(1508) & ChrW(1497) & ChrW(1497) & ChrW(1504) & ChrW(1497) & ChrW(1501) & "' -> " & ChrW(1489) & ChrW(1514) & ChrW(1495) & ChrW(1514) & ChrW(1497) & ChrW(1514) & " " & ChrW(1492) & ChrW(1495) & ChrW(1500) & ChrW(1493) & ChrW(1503) & " " & ChrW(1505) & ChrW(1502) & ChrW(1504) & ChrW(1493) & " V " & ChrW(1489) & ChrW(1514) & ChrW(1497) & ChrW(1489) & ChrW(1514) & " '" & ChrW(1489) & ChrW(1496) & ChrW(1500) & " " & ChrW(1495) & ChrW(1505) & ChrW(1497) & ChrW(1502) & ChrW(1492) & "' (Unblock) " & ChrW(1493) & ChrW(1500) & ChrW(1495) & ChrW(1510) & ChrW(1493) & " " & ChrW(1488) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1512) & "."
        .Range("B8").Value = "3. " & ChrW(1488) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1512) & " " & ChrW(1502) & ChrW(1488) & ChrW(1511) & ChrW(1512) & ChrW(1493) & ":"
        .Range("C8").Value = ChrW(1508) & ChrW(1514) & ChrW(1495) & ChrW(1493) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1510) & ". " & ChrW(1488) & ChrW(1501) & " " & ChrW(1502) & ChrW(1493) & ChrW(1508) & ChrW(1497) & ChrW(1506) & " " & ChrW(1508) & ChrW(1505) & " " & ChrW(1510) & ChrW(1492) & ChrW(1493) & ChrW(1489) & " " & ChrW(1500) & ChrW(1502) & ChrW(1506) & ChrW(1500) & ChrW(1492) & ", " & ChrW(1500) & ChrW(1495) & ChrW(1510) & ChrW(1493) & " " & ChrW(1506) & ChrW(1500) & " '" & ChrW(1492) & ChrW(1508) & ChrW(1493) & ChrW(1498) & " " & ChrW(1514) & ChrW(1493) & ChrW(1499) & ChrW(1503) & " " & ChrW(1500) & ChrW(1494) & ChrW(1502) & ChrW(1497) & ChrW(1503) & "' (Enable Content)."
        .Range("B9").Value = "4. " & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1514) & " " & ChrW(1494) & ChrW(1502) & ChrW(1504) & ChrW(1497) & ChrW(1501) & ":"
        .Range("C9").Value = ChrW(1489) & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1492) & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1497) & ", " & ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(1493) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1492) & ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505) & " " & ChrW(1493) & ChrW(1492) & ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1492) & ChrW(1504) & ChrW(1493) & ChrW(1499) & ChrW(1495) & ChrW(1497) & ChrW(1514) & " " & ChrW(1502) & ChrW(1492) & ChrW(1514) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1496) & ChrW(1497) & ChrW(1501) & " " & ChrW(1492) & ChrW(1504) & ChrW(1508) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1501) & "."
        
        .Range("B6:B9").Font.Bold = True
        .Range("B6:C9").Font.Size = 12
        .Range("B6:C9").RowHeight = 25
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 100
        
        ' Add Button
        Dim btnInstall As Shape
        For Each btnInstall In .Shapes
            btnInstall.Delete
        Next btnInstall
        Set btnInstall = .Shapes.AddShape(msoShapeRoundedRectangle, .Range("B12").Left, .Range("B12").Top, 300, 40)
        btnInstall.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1489) & ChrW(1504) & ChrW(1514) & ChrW(1497) & ", " & ChrW(1488) & ChrW(1500) & " " & ChrW(1514) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1500) & ChrW(1497) & " " & ChrW(1497) & ChrW(1493) & ChrW(1514) & ChrW(1512) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1494) & ChrW(1492)
        btnInstall.TextFrame2.TextRange.Font.Size = 14
        btnInstall.TextFrame2.TextRange.Font.Bold = msoTrue
        btnInstall.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        btnInstall.Fill.ForeColor.RGB = RGB(200, 50, 50)
        btnInstall.OnAction = "HideInstallInst"
        
        .Tab.Color = RGB(0, 150, 0)
    End With
    
    ' --- Setup Op Sheet ---
    With wsOp
        .Range("B2").Value = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & " " & ChrW(1500) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514)
        .Range("B2").Font.Size = 24
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Color = RGB(0, 51, 102)
        
        .Range("B4").Value = ChrW(1488) & ChrW(1497) & ChrW(1498) & " " & ChrW(1500) & ChrW(1506) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1506) & ChrW(1501) & " " & ChrW(1492) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1489) & ChrW(1497) & ChrW(1493) & ChrW(1501) & "-" & ChrW(1497) & ChrW(1493) & ChrW(1501) & ":"
        .Range("B4").Font.Size = 14
        .Range("B4").Font.Bold = True
        
        .Range("B6").Value = ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & " (" & ChrW(1492) & ChrW(1508) & ChrW(1488) & ChrW(1504) & ChrW(1500) & " " & ChrW(1492) & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1497) & "):"
        .Range("C6").Value = ChrW(1489) & ChrW(1510) & ChrW(1491) & " " & ChrW(1497) & ChrW(1502) & ChrW(1497) & ChrW(1503) & " " & ChrW(1514) & ChrW(1512) & ChrW(1488) & ChrW(1493) & " 7 " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & ChrW(1497) & ChrW(1501) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1502) & ChrW(1502) & ChrW(1493) & ChrW(1505) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1501) & ". " & ChrW(1506) & ChrW(1500) & ChrW(1497) & ChrW(1499) & ChrW(1501) & " " & ChrW(1500) & ChrW(1506) & ChrW(1489) & ChrW(1493) & ChrW(1512) & " " & ChrW(1506) & ChrW(1500) & ChrW(1497) & ChrW(1492) & ChrW(1501) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & ChrW(1492) & ChrW(1505) & ChrW(1491) & ChrW(1512) & " (1 " & ChrW(1506) & ChrW(1491) & " 7) " & ChrW(1499) & ChrW(1491) & ChrW(1497) & " " & ChrW(1500) & ChrW(1492) & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1501) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1492) & ChrW(1500) & ChrW(1497) & ChrW(1498) & "."
        .Range("B7").Value = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501) & ":"
        .Range("C7").Value = ChrW(1504) & ChrW(1497) & ChrW(1514) & ChrW(1503) & " " & ChrW(1500) & ChrW(1513) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1513) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1488) & ChrW(1493) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492) & " " & ChrW(1489) & ChrW(1499) & ChrW(1500) & " " & ChrW(1513) & ChrW(1500) & ChrW(1489) & " " & ChrW(1489) & ChrW(1496) & ChrW(1489) & ChrW(1500) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1514) & ChrW(1488) & ChrW(1497) & ChrW(1502) & ChrW(1492) & " " & ChrW(1489) & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1492) & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1497) & "."
        .Range("B8").Value = ChrW(1488) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & " " & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & ":"
        .Range("C8").Value = ChrW(1488) & ChrW(1501) & " " & ChrW(1502) & ChrW(1513) & ChrW(1492) & ChrW(1493) & " " & ChrW(1502) & ChrW(1513) & ChrW(1514) & ChrW(1489) & ChrW(1513) & " " & ChrW(1488) & ChrW(1493) & " " & ChrW(1513) & ChrW(1488) & ChrW(1514) & ChrW(1501) & " " & ChrW(1512) & ChrW(1493) & ChrW(1510) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1502) & ChrW(1495) & ChrW(1491) & ChrW(1513) & ", " & ChrW(1500) & ChrW(1495) & ChrW(1510) & ChrW(1493) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " '" & ChrW(1488) & ChrW(1508) & ChrW(1505) & " " & ChrW(1492) & ChrW(1499) & ChrW(1500) & " " & ChrW(1500) & ChrW(1502) & ChrW(1510) & ChrW(1489) & " " & ChrW(1492) & ChrW(1514) & ChrW(1495) & ChrW(1500) & ChrW(1514) & ChrW(1497) & "'."
        .Range("B9").Value = ChrW(1492) & ChrW(1491) & ChrW(1508) & ChrW(1505) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & ":"
        .Range("C9").Value = ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1493) & ChrW(1501) & " " & ChrW(1492) & ChrW(1514) & ChrW(1492) & ChrW(1500) & ChrW(1497) & ChrW(1498) & ", " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 7 " & ChrW(1497) & ChrW(1488) & ChrW(1508) & ChrW(1513) & ChrW(1512) & " " & ChrW(1500) & ChrW(1499) & ChrW(1501) & " " & ChrW(1500) & ChrW(1492) & ChrW(1491) & ChrW(1508) & ChrW(1497) & ChrW(1505) & " " & ChrW(1488) & ChrW(1493) & " " & ChrW(1500) & ChrW(1513) & ChrW(1502) & ChrW(1493) & ChrW(1512) & " " & ChrW(1499) & "-PDF " & ChrW(1488) & ChrW(1514) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1492) & ChrW(1514) & ChrW(1493) & ChrW(1510) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "."
        
        .Range("B6:B9").Font.Bold = True
        .Range("B6:C9").Font.Size = 12
        .Range("B6:C9").RowHeight = 25
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 100
        
        ' Add Button
        Dim btnOp As Shape
        For Each btnOp In .Shapes
            btnOp.Delete
        Next btnOp
        Set btnOp = .Shapes.AddShape(msoShapeRoundedRectangle, .Range("B12").Left, .Range("B12").Top, 300, 40)
        btnOp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1489) & ChrW(1504) & ChrW(1514) & ChrW(1497) & ", " & ChrW(1488) & ChrW(1500) & " " & ChrW(1514) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1500) & ChrW(1497) & " " & ChrW(1497) & ChrW(1493) & ChrW(1514) & ChrW(1512) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1494) & ChrW(1492)
        btnOp.TextFrame2.TextRange.Font.Size = 14
        btnOp.TextFrame2.TextRange.Font.Bold = msoTrue
        btnOp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        btnOp.Fill.ForeColor.RGB = RGB(200, 50, 50)
        btnOp.OnAction = "HideOpInst"
        
        .Tab.Color = RGB(0, 150, 0)
    End With
    
    MsgBoxU ChrW(1492) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1493) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation
    On Error GoTo 0
End Sub

Public Sub HideInstallInst()
    On Error Resume Next
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("AA1").Value = "YES"
    Dim sInstall As String
    sInstall = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
    ThisWorkbook.Worksheets(sInstall).Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
    On Error GoTo 0
End Sub

Public Sub HideOpInst()
    On Error Resume Next
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("AA2").Value = "YES"
    Dim sOp As String
    sOp = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
    ThisWorkbook.Worksheets(sOp).Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------
' LOGGER FOR DEBUGGING HANGS
' -------------------------------------------------------------------------"""

content = content.replace("""' -------------------------------------------------------------------------
' LOGGER FOR DEBUGGING HANGS
' -------------------------------------------------------------------------""", macros)


# Update version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_181"', 'Attribute VB_Name = "Goren_Claude_V2_182"')
content = content.replace('VERSION: V2.181', 'VERSION: V2.182')
content = content.replace('APP_VERSION As String = "2.181"', 'APP_VERSION As String = "2.182"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.182 correctly!")
