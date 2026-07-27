import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.085.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub BuildReview()" in line_strip:
        in_sub = True
    elif in_sub and "End Sub" in line_strip:
        break
    elif in_sub and "outIdx = outIdx + 1" in line_strip:
        for j in range(i-2, i+15):
            print(f"[{j+1}] {lines[j].strip()}")
        break

