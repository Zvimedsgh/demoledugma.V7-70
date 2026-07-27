import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.082.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Private Sub ProcessAndLoadDataToMemory(" in line_strip:
        in_func = True
    elif in_func and "End Sub" in line_strip:
        break
    elif in_func:
        if "For " in line_strip or "Do " in line_strip or "While " in line_strip or "Find(" in line_strip:
            print(f"[{i+1}] {line_strip}")
        if "Application." in line_strip:
            print(f"[{i+1}] {line_strip}")

