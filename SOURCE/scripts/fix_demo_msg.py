import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.159.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.160.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_159"', 'Attribute VB_Name = "Goren_Claude_V2_160"')
content = content.replace('VERSION: V2.159', 'VERSION: V2.160')
content = content.replace('APP_VERSION As String = "2.159"', 'APP_VERSION As String = "2.160"')

# Fix Demo message
old_msg_code = """
        Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("E19").Left, wsMain.Range("E19").Top, 500, 40)
        shpDemoMsg.Name = "shpDemoMsgText"
        3687 shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp " & ChrW(1500) & " " & ChrW(8211) & " 054-6677396"
"""

# Text: "להתקנה והדרכה שלח וואטסאפ לטלפון 054-6677396"
# וואטסאפ = 1493,1493,1488,1496,1505,1488,1508
# לטלפון = 1500,1496,1500,1508,1493,1503

new_msg_code = """
        Dim tLeft As Single, tWidth As Single
        tLeft = wsMain.Range("F19").Left
        tWidth = wsMain.Range("I19").Left + wsMain.Range("I19").Width - tLeft
        Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, tLeft, wsMain.Range("F19").Top, tWidth, 40)
        shpDemoMsg.Name = "shpDemoMsgText"
        3687 shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp  " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & " 054-6677396"
"""
# Above uses " WhatsApp  לטלפון 054-6677396" (with 2 spaces before Hebrew word for phone).
# That will 100% force the spacing to be right and RTL/LTR boundaries to be clear.

content = content.replace(old_msg_code.strip(), new_msg_code.strip())

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.160")
