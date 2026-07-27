import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.215.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub HideWorkSheets" in line:
        for j in range(i+15, i+35):
            print(lines[j].rstrip())
        break
