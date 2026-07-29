import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.253.bas', 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('APP_VERSION As String = "2.250"', 'APP_VERSION As String = "2.254"')
content = content.replace('VERSION: V2.253', 'VERSION: V2.254')
content = content.replace('Error in V2.253!', 'Error in V2.254!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_253"', 'Attribute VB_Name = "Goren_Claude_V2_254"')
content = content.replace('With wsMain.Range("C19:I19")', 'With wsMain.Range("D19:I19")')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.254.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 254')
