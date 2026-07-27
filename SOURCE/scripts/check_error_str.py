import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.071.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492)" in line:
        print(f"[{i+1}] {line.strip()}")

