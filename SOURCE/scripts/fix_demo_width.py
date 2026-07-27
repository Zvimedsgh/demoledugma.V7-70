import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.162.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.163.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(len(lines)):
    lines[i] = lines[i].replace('Attribute VB_Name = "Goren_Claude_V2_162"', 'Attribute VB_Name = "Goren_Claude_V2_163"')
    lines[i] = lines[i].replace('VERSION: V2.162', 'VERSION: V2.163')
    lines[i] = lines[i].replace('APP_VERSION As String = "2.162"', 'APP_VERSION As String = "2.163"')

# Fix Demo message formatting and width
for i, line in enumerate(lines):
    if 'Set shpDemoMsg = wsMain.Shapes.AddTextbox(' in line:
        # replace the AddTextbox line
        lines[i] = '    Dim tLeft As Single, tWidth As Single\n    tLeft = wsMain.Range("E19").Left\n    tWidth = wsMain.Range("K19").Left + wsMain.Range("K19").Width - tLeft\n    Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, tLeft, wsMain.Range("F19").Top, tWidth, 40)\n'
    if 'shpDemoMsg.TextFrame2.TextRange.Text =' in line:
        # "להתקנה והדרכה שלח WhatsApp לטלפון: 054-6677396"
        # 1500 = ל
        # 1496 = ט
        # 1500 = ל
        # 1508 = פ
        # 1493 = ו
        # 1503 = ן
        # We will use Right-To-Left Mark (8207) before the phone number to ensure correct rendering.
        lines[i] = '    shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & ": " & ChrW(8207) & "054-6677396"\n'

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Created V2.163 correctly!")
