import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.249.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.249', 'VERSION: V2.250')
content = content.replace('Error in V2.249!', 'Error in V2.250!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_249"', 'Attribute VB_Name = "Goren_Claude_V2_250"')
content = content.replace('APP_VERSION As String = "2.249"', 'APP_VERSION As String = "2.250"')
content = content.replace('APP_VERSION = "2.249"', 'APP_VERSION = "2.250"')

# Change zoom range
old_zoom = '4902 wsMain.Range("A1:K13").Select'
new_zoom = '4902 wsMain.Range("A1:L20").Select'
content = content.replace(old_zoom, new_zoom)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.250.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 250')
