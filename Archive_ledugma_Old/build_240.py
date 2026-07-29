import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.239.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.239', 'VERSION: V2.240')
content = content.replace('Error in V2.239!', 'Error in V2.240!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_239"', 'Attribute VB_Name = "Goren_Claude_V2_240"')
content = content.replace('APP_VERSION As String = "2.239"', 'APP_VERSION As String = "2.240"')
content = content.replace('APP_VERSION = "2.239"', 'APP_VERSION = "2.240"')

# Fix OpenInstallInstructions macro
old_macro = """Public Sub OpenInstallInstructions()
    On Error GoTo err_h
    Dim url As String
    url = "https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=Z8iVEV"
    CreateObject("WScript.Shell").Run url
    Exit Sub
err_h:
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & Err.Description, vbCritical
End Sub"""

new_macro = """Public Sub OpenInstallInstructions()
    On Error GoTo err_h
    Dim url As String
    url = "https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=Z8iVEV"
    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True
    Exit Sub
err_h:
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & Err.Description, vbCritical
End Sub"""

content = content.replace(old_macro, new_macro)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.240.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 240')
