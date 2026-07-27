import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub ApplyCorrectionsAndBuildReports" in line:
        in_func = True
    if in_func:
        if "MsgBoxU" in line and "ERR_HANDLER" not in line:
            for j in range(max(0, i-5), min(len(lines), i+15)):
                print(f"[{j+1}] {lines[j].strip()}")
            print("-" * 20)
        if "End Sub" in line:
            break

