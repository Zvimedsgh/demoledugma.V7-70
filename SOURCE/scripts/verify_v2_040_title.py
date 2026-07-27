import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.040_20260702_1507.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3735, 3750):
    print(f"[{i}] {lines[i].strip()}")
