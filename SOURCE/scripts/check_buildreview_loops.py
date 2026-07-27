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
    elif in_sub:
        if "For " in line_strip or "Do " in line_strip or "While " in line_strip or "Wend" in line_strip or "Next " in line_strip:
            print(f"[{i+1}] {line_strip}")

