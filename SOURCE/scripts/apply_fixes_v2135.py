import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.134.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Make the WhatsApp text look like: ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(8207) & ChrW(1500) & "-054-6677396"
old_wa1 = 'ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp" & ChrW(1500) & " - 054-6677396"'
new_wa1 = 'ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(8207) & ChrW(1500) & "-054-6677396"'
content = content.replace(old_wa1, new_wa1)

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_134"', 'Attribute VB_Name = "Goren_Claude_V2_135"')
content = content.replace('VERSION: V2.134', 'VERSION: V2.135')
content = content.replace('APP_VERSION As String = "2.134"', 'APP_VERSION As String = "2.135"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.135.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.135 created.")
