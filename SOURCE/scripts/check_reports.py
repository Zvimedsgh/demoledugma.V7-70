import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub ApplyCorrectionsAndBuildReports" in line:
        in_func = True
    if in_func:
        if i - lines.index("Public Sub ApplyCorrectionsAndBuildReports()\n") > 600: break
        if "wsOut.Name =" in line or "sheetName =" in line:
            print(f"[{i+1}] {lines[i].strip()}")
        if "End Sub" in line:
            in_func = False
            break

