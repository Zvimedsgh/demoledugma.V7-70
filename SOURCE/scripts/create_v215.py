import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.214.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.215.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

new_macro = """
' MACRO: SetupInstructionsSheet
Public Sub SetupInstructionsSheet()
    Dim wsName1 As String, wsName2 As String
    wsName1 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) ' הוראות_התקנה
    wsName2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) ' הוראות התקנה

    On Error Resume Next
    Dim wsInst As Worksheet
    Set wsInst = ThisWorkbook.Worksheets(wsName1)
    If wsInst Is Nothing Then Set wsInst = ThisWorkbook.Worksheets(wsName2)
    If wsInst Is Nothing Then Set wsInst = ThisWorkbook.Worksheets("Sheet2")
    On Error GoTo 0
    
    If wsInst Is Nothing Then
        MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514), vbExclamation
        Exit Sub
    End If
    
    ' Remove existing buttons to avoid duplicates
    Dim shp As Shape
    For Each shp In wsInst.Shapes
        If shp.Name = "btnHideInst" Then shp.Delete
    Next shp
    
    ' Create new button at top-left (near column A, row 2)
    Set shp = wsInst.Shapes.AddShape(msoShapeRoundedRectangle, wsInst.Range("A2").Left, wsInst.Range("A2").Top, 250, 40)
    shp.Name = "btnHideInst"
    shp.Fill.ForeColor.RGB = RGB(0, 100, 200) ' Nice blue color
    ' "הבנתי, אל תציג יותר גיליון זה"
    shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1489) & ChrW(1504) & ChrW(1514) & ChrW(1497) & ", " & ChrW(1488) & ChrW(1500) & " " & ChrW(1514) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1497) & ChrW(1493) & ChrW(1514) & ChrW(1512) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1494) & ChrW(1492)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "HideInstallationInstructions"
    
    MsgBoxU ChrW(1492) & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " " & ChrW(1492) & ChrW(1493) & ChrW(1505) & ChrW(1508) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation
End Sub
"""

if "Public Sub SetupInstructionsSheet" not in content:
    content += new_macro

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_214"', 'Attribute VB_Name = "Goren_Claude_V2_215"')
content = content.replace('VERSION: V2.214', 'VERSION: V2.215')
content = content.replace('APP_VERSION As String = "2.214"', 'APP_VERSION As String = "2.215"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.215")
