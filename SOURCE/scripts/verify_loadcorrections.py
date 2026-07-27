import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_fixed.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(1633, 1636):
    print(f"[{i}] {lines[i].strip()}")

for i in range(1710, 1740):
    print(f"[{i}] {lines[i].strip()}")
