import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.046.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3550, 3570):
    print(f"[{i}] {lines[i].strip()}")
