import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.071.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Cells(10, 1).Value" in line and "wsMsg" in line:
        print(f"[{i+1}] {line.strip()}")

