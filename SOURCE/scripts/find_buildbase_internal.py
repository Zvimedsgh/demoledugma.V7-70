import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(1782, 1800):
    print(f"[{i}] {lines[i].strip()}")
