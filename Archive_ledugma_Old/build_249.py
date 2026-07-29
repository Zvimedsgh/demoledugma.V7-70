import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.248.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.248', 'VERSION: V2.249')
content = content.replace('Error in V2.248!', 'Error in V2.249!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_248"', 'Attribute VB_Name = "Goren_Claude_V2_249"')
content = content.replace('APP_VERSION As String = "2.248"', 'APP_VERSION As String = "2.249"')
content = content.replace('APP_VERSION = "2.248"', 'APP_VERSION = "2.249"')

# Change zoom range from A1:M23 to A1:L13 (L is column 12, wait, let's use A1:K14 to be safe)
# Let's see what the old code is
old_zoom = '4902 wsMain.Range("A1:M23").Select'
new_zoom = '4902 wsMain.Range("A1:K13").Select'
content = content.replace(old_zoom, new_zoom)

# Remove row 4 from wsInstall
old_install = """        .Range("B4").Value = ChrW(1492) & ChrW(1490) & ChrW(1506) & ChrW(1514) & ChrW(1501) & " " & ChrW(1500) & ChrW(1502) & _
        ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1492) & ChrW(1491) & ChrW(1497) & ChrW(1493) & ChrW(1493) & _
        ChrW(1495) & " " & ChrW(1491) & ChrW(1512) & ChrW(1498) & " " & ChrW(1492) & ChrW(1491) & ChrW(1508) & ChrW(1491) & _
        ChrW(1508) & ChrW(1503) & ". " & ChrW(1499) & ChrW(1491) & ChrW(1497) & " " & ChrW(1500) & ChrW(1492) & ChrW(1508) & _
        ChrW(1506) & ChrW(1497) & ChrW(1500) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1506) & _
        ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1489) & ChrW(1510) & ChrW(1493) & ChrW(1512) & ChrW(1492) & " " & _
        ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1502) & ChrW(1495)"""

new_install = """        .Range("B4").Value = "" """
content = content.replace(old_install, new_install)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.249.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 249')
