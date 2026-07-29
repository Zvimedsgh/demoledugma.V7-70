import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.09.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "' Automatically open PowerPoint without prompt (as requested by user)" in line:
        skip = True
        
        # Inject the new ordered code
        new_lines.append("' Clean up presentation sheets automatically\n")
        new_lines.append("Application.DisplayAlerts = False\n")
        new_lines.append("' (REMOVED: User requested that result sheets are NOT deleted after presentation)\n")
        new_lines.append("Application.DisplayAlerts = True\n\n")
        
        new_lines.append("Application.EnableEvents = True\n")
        new_lines.append("Application.ScreenUpdating = True\n")
        new_lines.append("Application.DisplayAlerts = True\n\n")
        
        new_lines.append("' Show success message in Excel FIRST\n")
        new_lines.append("MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & \" \" & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & \" \" & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & \"!\" & vbCrLf & vbCrLf & ChrW(1500) & ChrW(1508) & ChrW(1514) & ChrW(1497) & ChrW(1495) & ChrW(1492) & \" \" & ChrW(1500) & ChrW(1495) & ChrW(1509) & \" \" & ChrW(1488) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1512), vbInformation\n\n")
        
        new_lines.append("' NOW maximize and show PowerPoint AFTER user clicks OK\n")
        new_lines.append("On Error Resume Next\n")
        new_lines.append("If Not ppApp Is Nothing And Not ppPres Is Nothing Then\n")
        new_lines.append("    ppApp.Visible = True\n")
        new_lines.append("    ppApp.WindowState = 3 ' Max\n")
        new_lines.append("    ppApp.Activate\n")
        new_lines.append("    AppActivate \"PowerPoint\"\n")
        new_lines.append("End If\n")
        new_lines.append("On Error GoTo ERR_HANDLER\n\n")
        
        new_lines.append("910     Set ppPres = Nothing\n")
        new_lines.append("920     Set ppApp = Nothing\n\n")
        
    elif "1060    Exit Sub" in line and skip:
        skip = False
        new_lines.append(line)
    elif not skip:
        new_lines.append(line)

content = "".join(new_lines)

content = content.replace('APP_VERSION As String = "3.09"', 'APP_VERSION As String = "3.10"')
content = content.replace('VERSION: V3.09', 'VERSION: V3.10')
content = content.replace('Error in V3.09!', 'Error in V3.10!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_09"', 'Attribute VB_Name = "Goren_Claude_V3_10"')

changelog = """' CHANGES IN 3.10:
'   - BUGFIX: Re-injected the correct presentation completion sequence (MsgBox first, then PowerPoint) after discovering the previous injection failed silently.
"""
content = content.replace("' CHANGES IN 3.09:", changelog + "' CHANGES IN 3.09:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.10.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.10')
