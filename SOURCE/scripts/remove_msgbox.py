import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.158.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

start_idx = -1
end_idx = -1

for i, line in enumerate(lines):
    if "Dim askMsg As String" in line and start_idx == -1:
        start_idx = i
    if "910     Set ppPres = Nothing" in line and start_idx != -1:
        end_idx = i
        break

if start_idx != -1 and end_idx != -1:
    new_lines = lines[:start_idx]
    
    new_code = """
        ' Automatically open PowerPoint without prompt (as requested by user)
        On Error Resume Next
        If Not ppApp Is Nothing And Not ppPres Is Nothing Then
            Application.WindowState = xlMinimized
            ppApp.Visible = True
            ppApp.WindowState = 2 ' Min
            ppApp.WindowState = 3 ' Max
            ppApp.Activate
        End If
        On Error GoTo ERR_HANDLER
"""
    new_lines.extend([new_code])
    new_lines.extend(lines[end_idx:])
    
    # Update version
    for i in range(len(new_lines)):
        new_lines[i] = new_lines[i].replace('Attribute VB_Name = "Goren_Claude_V2_158"', 'Attribute VB_Name = "Goren_Claude_V2_159"')
        new_lines[i] = new_lines[i].replace('VERSION: V2.158', 'VERSION: V2.159')
        new_lines[i] = new_lines[i].replace('APP_VERSION As String = "2.158"', 'APP_VERSION As String = "2.159"')
        
    with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.159.bas', 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    print("Created V2.159 with auto-open PowerPoint.")
else:
    print("Could not find the target block.")

