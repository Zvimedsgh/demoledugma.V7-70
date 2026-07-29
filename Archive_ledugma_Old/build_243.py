import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.242.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.242', 'VERSION: V2.243')
content = content.replace('Error in V2.242!', 'Error in V2.243!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_242"', 'Attribute VB_Name = "Goren_Claude_V2_243"')
content = content.replace('APP_VERSION As String = "2.242"', 'APP_VERSION As String = "2.243"')
content = content.replace('APP_VERSION = "2.242"', 'APP_VERSION = "2.243"')

start_str = 'Public Sub OpenManualSheet()'
start_idx = content.find(start_str)
end_idx = content.find('End Sub', start_idx) + 7

new_macro = """Public Sub OpenManualSheet()
    On Error GoTo err_h
    Dim pdfPath As String
    
    ' Check same folder first
    pdfPath = ThisWorkbook.Path & "\\MANUAL.PDF"
    If Dir(pdfPath) <> "" Then
        ThisWorkbook.FollowHyperlink Address:=pdfPath, NewWindow:=True
        Exit Sub
    End If
    
    ' Check OneDrive
    pdfPath = Environ("OneDriveCommercial")
    If pdfPath = "" Then pdfPath = Environ("OneDrive")
    If pdfPath = "" Then pdfPath = Environ("OneDriveConsumer")
    
    If pdfPath <> "" Then
        pdfPath = pdfPath & "\\MANUAL.PDF"
        If Dir(pdfPath) <> "" Then
            ThisWorkbook.FollowHyperlink Address:=pdfPath, NewWindow:=True
            Exit Sub
        End If
    End If
    
    MsgBox ChrW(1500) & ChrW(1488) & ChrW(32) & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(32) & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & ChrW(32) & "MANUAL.PDF", vbCritical
    Exit Sub
err_h:
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & Err.Description, vbCritical
End Sub"""

content = content[:start_idx] + new_macro + content[end_idx:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.243.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 243')
