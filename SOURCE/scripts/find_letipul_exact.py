import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)" in line:
        print(f"[{i}] {line.strip()}")
