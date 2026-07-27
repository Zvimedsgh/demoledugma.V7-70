import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.071.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for line in lines:
    if "Cells(10, 1).Value" in line:
        print(line.strip())

