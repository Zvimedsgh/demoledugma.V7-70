import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.247.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.247', 'VERSION: V2.248')
content = content.replace('Error in V2.247!', 'Error in V2.248!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_247"', 'Attribute VB_Name = "Goren_Claude_V2_248"')
content = content.replace('APP_VERSION As String = "2.247"', 'APP_VERSION As String = "2.248"')
content = content.replace('APP_VERSION = "2.247"', 'APP_VERSION = "2.248"')

old_code = """    ' ---- Navigate to A1 & Fit Screen ----
4900 wsMain.Activate
4902 wsMain.Range("A1:M23").Select
4905 ActiveWindow.Zoom = True
4910 Application.Goto wsMain.Range("F10")"""

new_code = """    ' ---- Navigate to A1 & Fit Screen ----
4900 wsMain.Activate
    ' Collapse Ribbon & Formula Bar for better laptop viewing
    On Error Resume Next
    Application.DisplayFormulaBar = False
    If Application.CommandBars.GetPressedMso("MinimizeRibbon") = False Then
        Application.CommandBars.ExecuteMso "MinimizeRibbon"
    End If
    On Error GoTo ERR_HANDLER
    
4902 wsMain.Range("A1:M23").Select
4905 ActiveWindow.Zoom = True
    ' Force scroll to top-left to prevent cutoff
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
4910 Application.Goto wsMain.Range("F10")"""

content = content.replace(old_code, new_code)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.248.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 248')
