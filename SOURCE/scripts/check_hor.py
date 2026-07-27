import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "הוראות" in line:
        print(line.strip())
