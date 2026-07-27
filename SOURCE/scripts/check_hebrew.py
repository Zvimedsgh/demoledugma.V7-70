import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.084.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Function HebrewToKey" in line_strip:
        in_sub = True
    elif in_sub and "End Function" in line_strip:
        for j in range(i-30, i+2):
            print(f"[{j+1}] {lines[j].strip()}")
        break
    elif in_sub and "While" in line_strip:
        print(f"[{i+1}] {line_strip}")
    elif in_sub and "For " in line_strip:
        print(f"[{i+1}] {line_strip}")

