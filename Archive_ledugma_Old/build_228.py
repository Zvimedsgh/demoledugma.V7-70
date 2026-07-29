import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.227.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.227', 'VERSION: V2.228')
content = content.replace('Error in V2.227!', 'Error in V2.228!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_227"', 'Attribute VB_Name = "Goren_Claude_V2_228"')
content = content.replace('APP_VERSION = "2.226"', 'APP_VERSION = "2.228"')
content = content.replace('APP_VERSION = "2.227"', 'APP_VERSION = "2.228"')

# Fix white cells in ApplyDemoLockOnOpen
content = content.replace('wsMain.Range("G3:G4").Interior.ColorIndex = xlNone', 'wsMain.Range("G3:G4").Interior.Color = RGB(255, 245, 230)')

# Add version update to ApplyDemoLockOnOpen
version_str = 'wsMain.Range("A22").Value = ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " v" & APP_VERSION\n    wsMain.Range("A22").Font.Size = 10\n    wsMain.Range("A22").Font.Color = RGB(150, 150, 150)'
content = content.replace('ThisWorkbook.Worksheets(homeSheetName).Activate', version_str + '\n    ThisWorkbook.Worksheets(homeSheetName).Activate')

old_macro = """Public Sub OpenManualSheet()
    On Error Resume Next
    Dim sOp As String
    sOp = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
    ThisWorkbook.Worksheets(sOp).Activate
    On Error GoTo 0
End Sub"""

new_macro = """Public Sub OpenManualSheet()
    On Error GoTo err_h
    Dim sOp As String
    sOp = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sOp)
    
    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If
    
    ws.Activate
    Exit Sub
    
err_h:
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & Err.Description, vbCritical
End Sub"""

content = content.replace(old_macro, new_macro)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.228.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 228')
