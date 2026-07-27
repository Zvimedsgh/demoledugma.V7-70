import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.052_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_func = False
line_num = 10

for line in lines:
    if "Private Sub ExportCompCharts" in line:
        in_func = True
        line_num = 10
        new_lines.append(line)
        continue
    if in_func and "End Sub" in line:
        in_func = False
        new_lines.append(line)
        continue
        
    if in_func:
        # Strip existing line numbers (10 to 9999 followed by space)
        stripped = line.lstrip()
        m = re.match(r'^(\d+)\s+(.*)', stripped)
        if m:
            stripped = m.group(2)
            
        if stripped != "" and not stripped.startswith("'") and not stripped.startswith("On Error") and not stripped.endswith(":") and "Dim " not in stripped and "ReDim " not in stripped:
            new_lines.append(f"{line_num} " + line.lstrip(' 0123456789'))
            line_num += 10
        else:
            new_lines.append(line)
    else:
        new_lines.append(line)

# Let's also bump version
for i, line in enumerate(new_lines):
    if "Attribute VB_Name =" in line:
        new_lines[i] = line.replace("V2_052", "V2_053")
    if "Private Const APP_VERSION As String =" in line:
        new_lines[i] = line.replace("2.052", "2.053")
    if "VERSION: V2.052" in line:
        new_lines[i] = line.replace("2.052", "2.053")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.053_LNUM.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.053_LNUM")
