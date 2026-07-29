import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.244.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.244', 'VERSION: V2.245')
content = content.replace('Error in V2.244!', 'Error in V2.245!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_244"', 'Attribute VB_Name = "Goren_Claude_V2_245"')
content = content.replace('APP_VERSION As String = "2.244"', 'APP_VERSION As String = "2.245"')
content = content.replace('APP_VERSION = "2.244"', 'APP_VERSION = "2.245"')

start_str = 'Public Sub OpenManualSheet()'
start_idx = content.find(start_str)
end_idx = content.find('End Sub', start_idx) + 7

new_macro = """Public Sub OpenManualSheet()
    On Error GoTo err_h
    Dim url As String
    url = "https://gorentec-my.sharepoint.com/:b:/g/personal/zvi_gorentech_co_il/IQBt2Ms0oWLySpRBl-6OY68MAXnBN2jyzpD3y43ZHmIUBwU?e=h8a5ZG"
    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True
    Exit Sub
err_h:
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & Err.Description, vbCritical
End Sub"""

content = content[:start_idx] + new_macro + content[end_idx:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.245.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 245')
