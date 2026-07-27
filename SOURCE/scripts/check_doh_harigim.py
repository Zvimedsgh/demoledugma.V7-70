import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.072.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ChrW(1491) & ChrW(1493) & ChrW(1495) & \" \" & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501)" in line:
        print(f"[{i+1}] {line.strip()}")

