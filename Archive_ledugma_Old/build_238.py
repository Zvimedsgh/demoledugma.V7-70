import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.236.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.236', 'VERSION: V2.238')
content = content.replace('Error in V2.236!', 'Error in V2.238!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_236"', 'Attribute VB_Name = "Goren_Claude_V2_238"')
content = content.replace('APP_VERSION As String = "2.236"', 'APP_VERSION As String = "2.238"')
content = content.replace('APP_VERSION = "2.236"', 'APP_VERSION = "2.238"')

start_str = 'Public Sub OpenManualSheet()'
start_idx = content.find(start_str)
end_idx = content.find('End Sub', start_idx) + 7

new_macro = """Public Sub OpenManualSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Dim foundWs As Worksheet
    Dim allNames As String
    
    Set foundWs = ThisWorkbook.Worksheets(ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500))
    
    If foundWs Is Nothing Then
        For Each ws In ThisWorkbook.Worksheets
            allNames = allNames & ws.Name & vbCrLf
            If InStr(ws.Name, ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)) > 0 Then
                Set foundWs = ws
                Exit For
            End If
            If InStr(ws.Name, ChrW(1502) & ChrW(1491) & ChrW(1512) & ChrW(1497) & ChrW(1498)) > 0 Then
                Set foundWs = ws
                Exit For
            End If
        Next ws
    End If
    
    If Not foundWs Is Nothing Then
        foundWs.Visible = xlSheetVisible
        foundWs.Activate
    Else
        MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & vbCrLf & vbCrLf & "Available:" & vbCrLf & allNames, vbCritical
    End If
    On Error GoTo 0
End Sub"""

content = content[:start_idx] + new_macro + content[end_idx:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.238.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 238')
