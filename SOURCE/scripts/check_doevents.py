import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.084.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub BuildReview()" in line:
        pass
    if "DoEvents" in line:
        print(f"[{i+1}] {line.strip()}")

