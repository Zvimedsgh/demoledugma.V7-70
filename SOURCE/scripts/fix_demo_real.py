import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.160.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.161.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(len(lines)):
    lines[i] = lines[i].replace('Attribute VB_Name = "Goren_Claude_V2_160"', 'Attribute VB_Name = "Goren_Claude_V2_161"')
    lines[i] = lines[i].replace('VERSION: V2.160', 'VERSION: V2.161')
    lines[i] = lines[i].replace('APP_VERSION As String = "2.160"', 'APP_VERSION As String = "2.161"')

# Now find the block
for i, line in enumerate(lines):
    if 'Set shpDemoMsg = wsMain.Shapes.AddTextbox(' in line:
        # replace the AddTextbox line
        lines[i] = '    Dim tLeft As Single, tWidth As Single\n    tLeft = wsMain.Range("F19").Left\n    tWidth = wsMain.Range("I19").Left + wsMain.Range("I19").Width - tLeft\n    Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, tLeft, wsMain.Range("F19").Top, tWidth, 40)\n'
    if 'shpDemoMsg.TextFrame2.TextRange.Text =' in line:
        lines[i] = '    shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp   " & ChrW(1500) & " - 054-6677396"\n'

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Created V2.161 correctly!")
