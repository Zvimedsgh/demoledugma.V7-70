import sys
import time

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.084.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub BuildReview()" in line_strip:
        in_sub = True
    elif in_sub and "End Sub" in line_strip:
        break
    elif in_sub and "Dim " in line_strip and " As " in line_strip:
        pass

