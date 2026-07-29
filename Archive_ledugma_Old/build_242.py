import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.241.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.241', 'VERSION: V2.242')
content = content.replace('Error in V2.241!', 'Error in V2.242!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_241"', 'Attribute VB_Name = "Goren_Claude_V2_242"')
content = content.replace('APP_VERSION As String = "2.241"', 'APP_VERSION As String = "2.242"')
content = content.replace('APP_VERSION = "2.241"', 'APP_VERSION = "2.242"')

# Clean up the MsgBox for missing manual sheet
old_msg = 'MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & vbCrLf & vbCrLf & "Available:" & vbCrLf & allNames, vbCritical'
new_msg = 'MsgBox ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & ChrW(32) & ChrW(1502) & ChrW(1491) & ChrW(1512) & ChrW(1497) & ChrW(1499) & ChrW(32) & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(32) & ChrW(1496) & ChrW(1512) & ChrW(1501) & ChrW(32) & ChrW(1492) & ChrW(1493) & ChrW(1511) & ChrW(1501) & ChrW(32) & ChrW(1489) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & ChrW(46), vbInformation'
content = content.replace(old_msg, new_msg)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.242.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 242')
