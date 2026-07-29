import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.277.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "' 10 - Toggle Layout" in line:
        skip = True
        continue
    
    if skip and 'shpToggleLayout.OnAction' in line:
        skip = False
        continue
        
    if skip:
        continue

    # Right before A00 ends, add the auto-zoom
    if 'Application.ScreenUpdating = True' in line and '4040' in lines[i-1]:
        new_lines.append("    ' Auto-detect screen width and set zoom\n")
        new_lines.append("    If Application.UsableWidth < 1200 Then\n")
        new_lines.append("        ActiveWindow.Zoom = 75\n")
        new_lines.append("    Else\n")
        new_lines.append("        ActiveWindow.Zoom = 100\n")
        new_lines.append("    End If\n")
        new_lines.append(line)
        continue

    # Delete the ToggleDisplayMode sub
    if 'Public Sub ToggleDisplayMode()' in line:
        skip = True
        continue
        
    if skip and 'End Sub' in line:
        skip = False
        continue

    if skip:
        continue

    # Move Exit button up
    if 'shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 95.9, 385.5, 151.1, 35.3)' in line:
        line = line.replace('385.5', '265.1')
        
    new_lines.append(line)

content = "".join(new_lines)
content = content.replace('APP_VERSION As String = "2.277"', 'APP_VERSION As String = "2.278"')
content = content.replace('VERSION: V2.277', 'VERSION: V2.278')
content = content.replace('Error in V2.277!', 'Error in V2.278!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_277"', 'Attribute VB_Name = "Goren_Claude_V2_278"')

changelog = """' CHANGES IN 2.278:
'   - UI: Removed "Laptop View" toggle button entirely.
'   - UI: Added automatic screen width detection (Zoom=75% if UsableWidth < 1200).
"""
content = content.replace("' CHANGES IN 2.277:", changelog + "' CHANGES IN 2.277:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.278.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 278')
