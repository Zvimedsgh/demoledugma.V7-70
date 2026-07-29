import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.08.bas', 'r', encoding='utf-8') as f:
    content = f.read()

target_old = "615     ppApp.Visible = True\n620     Set ppPres = ppApp.Presentations.Add"

target_new = """615     ppApp.Visible = True
        On Error Resume Next
        ppApp.WindowState = 2 ' Minimized for speed and to keep Excel in focus
        On Error GoTo ERR_HANDLER
620     Set ppPres = ppApp.Presentations.Add"""

content = content.replace(target_old, target_new)


content = content.replace('APP_VERSION As String = "3.08"', 'APP_VERSION As String = "3.09"')
content = content.replace('VERSION: V3.08', 'VERSION: V3.09')
content = content.replace('Error in V3.08!', 'Error in V3.09!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_08"', 'Attribute VB_Name = "Goren_Claude_V3_09"')

changelog = """' CHANGES IN 3.09:
'   - BUGFIX: Fixed the missing 'WindowState = 2' (minimize PowerPoint) command that failed to inject previously. Now PowerPoint will properly minimize itself while building, keeping Excel in the foreground so the user sees the work and the final message.
"""
content = content.replace("' CHANGES IN 3.08:", changelog + "' CHANGES IN 3.08:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.09.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.09')
