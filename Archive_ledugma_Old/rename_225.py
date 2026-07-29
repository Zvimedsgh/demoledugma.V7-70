import sys
with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.225.bas', 'r', encoding='utf-8') as f:
    content = f.read()
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_224"', 'Attribute VB_Name = "Goren_Claude_V2_225"')
content = content.replace('VERSION: V2.224', 'VERSION: V2.225')
with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.225.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Updated 225 name.')
