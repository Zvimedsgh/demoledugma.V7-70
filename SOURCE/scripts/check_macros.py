import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.060.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub " in line_strip:
        print(line_strip)

