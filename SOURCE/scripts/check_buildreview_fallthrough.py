import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.071.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub BuildReview()" in line_strip:
        in_func = True
    elif in_func and "End Sub" in line_strip:
        break
    elif in_func and "ERR_HANDLER:" in line_strip:
        for j in range(i-10, i):
            print(f"[{j+1}] {lines[j].strip()}")

