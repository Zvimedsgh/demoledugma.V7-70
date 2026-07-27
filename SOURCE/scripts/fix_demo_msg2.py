import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.160.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_str = 'ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp  " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & " 054-6677396"'
new_str = 'ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp   " & ChrW(1500) & " - 054-6677396"'

content = content.replace(old_str, new_str)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Updated spacing in V2.160")
