import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if chr(1500) + chr(1496) + chr(1497) + chr(1508) + chr(1493) + chr(1500) in line:
        print(f"[{i}] {line.strip()}")
