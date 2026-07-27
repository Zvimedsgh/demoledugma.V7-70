import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.097.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub BuildReview()" in line:
        in_func = True
    elif in_func and "End Sub" in line:
        break
    elif in_func and "srcData =" in line:
        for j in range(i-5, min(len(lines), i+5)):
            print(f"[{j+1}] {lines[j].strip()}")
        break

