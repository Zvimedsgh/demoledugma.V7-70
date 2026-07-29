import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.246.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.246', 'VERSION: V2.247')
content = content.replace('Error in V2.246!', 'Error in V2.247!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_246"', 'Attribute VB_Name = "Goren_Claude_V2_247"')
content = content.replace('APP_VERSION As String = "2.246"', 'APP_VERSION As String = "2.247"')
content = content.replace('APP_VERSION = "2.246"', 'APP_VERSION = "2.247"')

old_code = """    ' ---- Navigate to A1 ----
4900 wsMain.Activate
4910 Application.Goto wsMain.Range("F10")"""

new_code = """    ' ---- Navigate to A1 & Fit Screen ----
4900 wsMain.Activate
4902 wsMain.Range("A1:M23").Select
4905 ActiveWindow.Zoom = True
4910 Application.Goto wsMain.Range("F10")"""

content = content.replace(old_code, new_code)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.247.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 247')
