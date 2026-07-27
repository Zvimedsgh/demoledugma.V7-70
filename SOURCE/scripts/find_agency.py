import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "AgencyName" in line or chr(1513) + chr(1501) + " " + chr(1505) + chr(1493) + chr(1499) + chr(1504) + chr(1493) + chr(1514) in line:
        print(f"[{i}] {line.strip()}")
