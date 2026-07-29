import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.256.bas', 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('APP_VERSION As String = "2.256"', 'APP_VERSION As String = "2.257"')
content = content.replace('VERSION: V2.256', 'VERSION: V2.257')
content = content.replace('Error in V2.256!', 'Error in V2.257!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_256"', 'Attribute VB_Name = "Goren_Claude_V2_257"')

# Update Changelog
changelog = """' CHANGES IN 2.257:
'   - UI: Fixed zoom on laptops by selecting A1:P22 before Zoom = True to force zoom out.
"""
content = content.replace("' CHANGES IN 2.256:", changelog + "' CHANGES IN 2.256:")

# Fix the zoom
content = content.replace('wsMain.Range("A1:L20").Select', 'wsMain.Range("A1:P22").Select')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.257.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 257')
