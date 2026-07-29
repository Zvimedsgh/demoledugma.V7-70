import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.243.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.243', 'VERSION: V2.244')
content = content.replace('Error in V2.243!', 'Error in V2.244!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_243"', 'Attribute VB_Name = "Goren_Claude_V2_244"')
content = content.replace('APP_VERSION As String = "2.243"', 'APP_VERSION As String = "2.244"')
content = content.replace('APP_VERSION = "2.243"', 'APP_VERSION = "2.244"')

start_str = 'Public Sub OpenManualSheet()'
start_idx = content.find(start_str)
end_idx = content.find('End Sub', start_idx) + 7

new_macro = """Public Sub OpenManualSheet()
    On Error Resume Next
    Dim pdfPath As String
    
    ' 1. Try relative to the workbook
    pdfPath = ThisWorkbook.Path
    If Left(pdfPath, 4) = "http" Then
        pdfPath = pdfPath & "/MANUAL.PDF"
    Else
        pdfPath = pdfPath & "\\MANUAL.PDF"
    End If
    
    ThisWorkbook.FollowHyperlink Address:=pdfPath, NewWindow:=True
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    
    ' 2. Try OneDrive environment variables if the above failed
    Dim envPath As String
    envPath = Environ("OneDriveCommercial")
    If envPath = "" Then envPath = Environ("OneDrive")
    If envPath = "" Then envPath = Environ("OneDriveConsumer")
    
    If envPath <> "" Then
        pdfPath = envPath & "\\MANUAL.PDF"
        ThisWorkbook.FollowHyperlink Address:=pdfPath, NewWindow:=True
        If Err.Number = 0 Then Exit Sub
        Err.Clear
    End If
    
    ' If all failed, show error
    MsgBox ChrW(1500) & ChrW(1488) & ChrW(32) & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(32) & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & ChrW(32) & "MANUAL.PDF", vbCritical
End Sub"""

content = content[:start_idx] + new_macro + content[end_idx:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.244.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 244')
