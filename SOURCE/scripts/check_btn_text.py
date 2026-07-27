import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.172.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "shp.TextFrame2.TextRange.Text =" in line and "shp.Name = " in lines[i-3]:
        print(lines[i-3].strip())
        print(line.strip())

