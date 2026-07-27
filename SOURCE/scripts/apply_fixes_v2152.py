import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.151.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target = """Call HideWorkSheets
2795    MsgBoxU ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501), vbInformation"""

new_code = """UpdateClientList
Call ShowHiddenSheets
2795    MsgBoxU ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501), vbInformation"""

if target in content:
    content = content.replace(target, new_code)
    print("Fixed ApplyCorrectionsAndBuildReports - added UpdateClientList and ShowHiddenSheets.")
else:
    print("Could not find target string.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_151"', 'Attribute VB_Name = "Goren_Claude_V2_152"')
content = content.replace('VERSION: V2.151', 'VERSION: V2.152')
content = content.replace('APP_VERSION As String = "2.151"', 'APP_VERSION As String = "2.152"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.152.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.152 created.")
