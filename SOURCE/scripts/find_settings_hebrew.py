import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "הגדרות" in line or "ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)" in line:
        print(f"[{i}] {lines[i].strip()}")
