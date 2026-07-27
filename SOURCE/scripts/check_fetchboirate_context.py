import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.092.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3650, 3700):
    print(f"[{i+1}] {lines[i].strip()}")

