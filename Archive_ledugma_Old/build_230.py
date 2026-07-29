import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.229.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.229', 'VERSION: V2.230')
content = content.replace('Error in V2.229!', 'Error in V2.230!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_229"', 'Attribute VB_Name = "Goren_Claude_V2_230"')

# Fix APP_VERSION constant
content = content.replace('APP_VERSION As String = "2.221"', 'APP_VERSION As String = "2.230"')

# Change msoShapeDiamond to msoShapeOval
content = content.replace('msoShapeDiamond,', 'msoShapeOval,')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.230.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 230')
