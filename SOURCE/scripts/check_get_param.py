import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.078.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Function GetStringParameter" in line_strip:
        in_sub = True
    elif in_sub and "End Function" in line_strip:
        break
    elif in_sub and "1710 " in line_strip:
        print(f"[{i+1}] {lines[i].strip()}")

