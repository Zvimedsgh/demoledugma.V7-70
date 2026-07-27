import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

macro = """
Public Sub DemoModeRestricted()
    MsgBoxU ChrW(1492) & ChrW(1488) & ChrW(1508) & ChrW(1513) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1494) & ChrW(1502) & ChrW(1497) & ChrW(1504) & ChrW(1492) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & ".", vbInformation, ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1491) & ChrW(1502) & ChrW(1493)
End Sub
"""

if "Public Sub DemoModeRestricted" not in content:
    content += "\n" + macro
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Added DemoModeRestricted macro")
else:
    print("Macro already exists")
