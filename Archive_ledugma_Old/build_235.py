import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.234.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.234', 'VERSION: V2.235')
content = content.replace('Error in V2.234!', 'Error in V2.235!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_234"', 'Attribute VB_Name = "Goren_Claude_V2_235"')
content = content.replace('APP_VERSION As String = "2.234"', 'APP_VERSION As String = "2.235"')

# Remove Hyperlinks.Add and add OnAction for shpInstallMsg
old_hl = 'wsMain.Hyperlinks.Add Anchor:=shpInstall, Address:="https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=Z8iVEV"'
new_hl = 'shpInstall.OnAction = "OpenInstallInstructions"'
content = content.replace(old_hl, new_hl)

# Add OpenInstallInstructions macro
new_macro = """
Public Sub OpenInstallInstructions()
    On Error GoTo err_h
    Dim url As String
    url = "https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=Z8iVEV"
    CreateObject("WScript.Shell").Run """" & url & """"
    Exit Sub
err_h:
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & Err.Description, vbCritical
End Sub
"""

content = content + new_macro

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.235.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 235')
