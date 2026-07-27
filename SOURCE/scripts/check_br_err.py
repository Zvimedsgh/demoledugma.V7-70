import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.198.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_br = False
for i, line in enumerate(lines):
    if "Public Sub BuildReview()" in line:
        in_br = True
    if in_br and ("Err.Raise" in line or "1004" in line):
        print(f"[{i+1}] {line.strip()}")
    if in_br and "End Sub" in line:
        break

