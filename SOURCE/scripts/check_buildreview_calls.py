import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.082.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub BuildReview()" in line_strip:
        in_sub = True
    elif in_sub and "End Sub" in line_strip:
        break
    elif in_sub and "Call " in line_strip:
        print(f"[{i+1}] {line_strip}")
    elif in_sub and "Build" in line_strip and not "BuildReview" in line_strip:
        print(f"[{i+1}] {line_strip}")
    elif in_sub and "Process" in line_strip:
        print(f"[{i+1}] {line_strip}")
    elif in_sub and "Apply" in line_strip:
        print(f"[{i+1}] {line_strip}")

