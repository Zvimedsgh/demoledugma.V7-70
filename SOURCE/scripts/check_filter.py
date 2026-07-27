import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.195.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_filter = False
for i, line in enumerate(lines):
    if "Private Sub FilterData" in line:
        in_filter = True
    if in_filter and "srcData" in line and "RAW_" in line:
        print(f"[{i+1}] {line.strip()}")
    if in_filter and "End Sub" in line:
        break

