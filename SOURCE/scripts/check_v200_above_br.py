import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.200.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_br = False
for i, line in enumerate(lines):
    if "Public Sub BuildReview()" in line:
        for j in range(i-55, i):
            print(f"[{j+1}] {lines[j].strip()}")
        break
