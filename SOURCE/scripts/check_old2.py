import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub ApplyCorrectionsAndBuildReports(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 500: break
        if "threshold" in line or "GoTo NextSrcRow" in line:
            print(f"[{i+1}] {line.strip()}")
        if "End Sub" in line:
            break

