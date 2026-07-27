import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.049_20260702_1635.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(850, 890):
    print(f"[{i}] {lines[i].strip()}")
