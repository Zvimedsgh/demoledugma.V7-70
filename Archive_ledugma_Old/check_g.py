import sys

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.252.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()
for i, line in enumerate(lines):
    if 'wsMain.Range("G' in line and 'Merge' in line:
        print(f'{i}: {line.strip()}')
