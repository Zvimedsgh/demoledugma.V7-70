with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# We will replace the text assignment in shpInstall.TextFrame2.TextRange
old_text_block = r'\.Text = ChrW\(1512\).*?"6"'
new_text_block = r'''
.Text = ChrW(1500) & ChrW(1510) & ChrW(1508) & ChrW(1497) & ChrW(1497) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1489) & ChrW(1488) & ChrW(1497) & ChrW(1504) & ChrW(1496) & ChrW(1512) & ChrW(1504) & ChrW(1496) & " - " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1499) & ChrW(1488) & ChrW(1503) & vbCrLf & _
ChrW(1500) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & " 054-6677396"
'''.strip()

text = re.sub(old_text_block, new_text_block, text, flags=re.DOTALL)

# Add OnAction to shpInstall
if 'shpInstall.OnAction =' not in text:
    text = text.replace('shpInstall.TextFrame2.VerticalAnchor = msoAnchorMiddle', 'shpInstall.TextFrame2.VerticalAnchor = msoAnchorMiddle\nshpInstall.OnAction = "OpenInstallWeb"')

# Add OpenInstallWeb macro
open_install_web_macro = """
Public Sub OpenInstallWeb()
    On Error Resume Next
    ' Replace this URL with the actual URL for the installation instructions
    ActiveWorkbook.FollowHyperlink "https://gorentec-my.sharepoint.com/"
    On Error GoTo 0
End Sub
"""
text += "\n" + open_install_web_macro

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated installation shape to be a web link!")
