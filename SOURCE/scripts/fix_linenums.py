import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.058.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        lines[i] = 'Attribute VB_Name = "Goren_Claude_V2_058"\n'
    elif "VERSION: " in line:
        lines[i] = "' VERSION: V2.058\n"
    elif "Private Const APP_VERSION As String =" in line:
        lines[i] = 'Private Const APP_VERSION As String = "2.058"\n'

    if "Private Sub ExportCompCharts" in line or "Private Sub ExportTotalChart" in line:
        in_func = True
        
    if in_func:
        # Strip line number
        m = re.match(r'^(\s*)\d+\s+(.*)', line)
        if m:
            lines[i] = m.group(1) + m.group(2) + "\n"
        
        if "End Sub" in line:
            in_func = False

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Created V2.058")

