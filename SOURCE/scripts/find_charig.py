import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if chr(1495) + chr(1512) + chr(1497) + chr(1490) in line:
        print(f"[{i}] {line.strip()}")
