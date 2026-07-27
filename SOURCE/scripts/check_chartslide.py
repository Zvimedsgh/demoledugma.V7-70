import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.109.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Sub BuildChartSlide" in line:
        in_func = True
        print(f"[{i+1}] {line.strip()}")
    elif in_func and "End Sub" in line:
        print(f"[{i+1}] {line.strip()}")
        break
    elif in_func:
        if "AddTextbox" in line or "Top =" in line or "Left =" in line:
            print(f"[{i+1}] {line.strip()}")

