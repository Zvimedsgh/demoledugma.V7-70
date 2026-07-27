import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.151.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
found = False
for i, line in enumerate(lines):
    if "Public Sub ApplyCorrectionsAndBuildReports(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 2000: break
        if "UpdateClientList" in line:
            print(f"[{i+1}] {line.strip()}")
            found = True
        if "End Sub" in line:
            break

if not found:
    print("UpdateClientList NOT FOUND in ApplyCorrectionsAndBuildReports")

