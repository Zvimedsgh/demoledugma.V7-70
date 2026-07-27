import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if chr(1500) + chr(1496) + chr(1497) + chr(1508) in line: # "לטיפ"
        print(f"[{i}] {line.strip()}")
