import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.180.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511)" in line:
        print(f"[{i-2}] {lines[i-2].strip()}")
        print(f"[{i-1}] {lines[i-1].strip()}")
        print(f"[{i}] {line.strip()}")

