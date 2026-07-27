import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.135.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update WhatsApp text
old_text = 'ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(8207) & ChrW(1500) & "-054-6677396"'
new_text = 'ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp " & ChrW(1500) & " " & ChrW(8211) & " 054-6677396"'
content = content.replace(old_text, new_text)

# 2. Update Demo message position (move from F19 to E19 so it's centered)
content = content.replace('wsMain.Range("F19").Left, wsMain.Range("F19").Top', 'wsMain.Range("E19").Left, wsMain.Range("E19").Top')


# Update Version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_135"', 'Attribute VB_Name = "Goren_Claude_V2_136"')
content = content.replace('VERSION: V2.135', 'VERSION: V2.136')
content = content.replace('APP_VERSION As String = "2.135"', 'APP_VERSION As String = "2.136"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.136.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.136 created.")

