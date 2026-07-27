import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
for i, line in enumerate(lines):
    if "Public Sub Create" in line or "Public Sub Setup" in line:
        in_sub = True
    if in_sub and "COMPANY" in line and "wsSettings.Cells" in line:
        for j in range(i-5, i+20):
            print(f"[{j+1}] {lines[j].strip()}")
        break

