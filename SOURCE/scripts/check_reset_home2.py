import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.124.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "G10" in line and "RGB(220, 240, 220)" in line:
        for j in range(max(0, i-2), min(len(lines), i+2)):
            print(f"[{j+1}] {lines[j].strip()}")

