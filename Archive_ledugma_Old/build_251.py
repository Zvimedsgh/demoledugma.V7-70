import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.250.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for line in lines:
    if line.startswith('VERSION: V2.250'):
        new_lines.append(line.replace('250', '251'))
    elif line.startswith('APP_VERSION As String'):
        new_lines.append(line.replace('250', '251'))
    elif line.startswith('APP_VERSION = '):
        new_lines.append(line.replace('250', '251'))
    elif 'Error in V2.250' in line:
        new_lines.append(line.replace('250', '251'))
    elif 'Attribute VB_Name = "Goren_Claude_V2_250"' in line:
        new_lines.append(line.replace('250', '251'))
    
    # Remove shpInstallMsg logic
    elif 'If hl.Shape.Name = "shpInstallMsg" Then hl.Delete' in line:
        new_lines.append("        ' " + line.lstrip())
    elif 'wsMain.Shapes("shpInstallMsg").Delete' in line:
        new_lines.append("    ' " + line.lstrip())
    elif "' --- Add Right Button: Installation Instructions" in line:
        skip = True
        new_lines.append("        ' Removed shpInstallMsg\n")
    elif skip and "shpInstall.TextFrame2.VerticalAnchor = msoAnchorMiddle" in line:
        skip = False
    elif skip:
        pass
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.251.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)
print('Generated 251')
