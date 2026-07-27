import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_fixed.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(945, 975):
    print(f"[{i}] {lines[i].strip()}")
