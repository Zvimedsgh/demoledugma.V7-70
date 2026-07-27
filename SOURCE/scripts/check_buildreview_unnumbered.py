import sys
import re

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
    elif in_func:
        if line_strip == "" or line_strip.startswith("'"):
            continue
        if not re.match(r'^\d+\s', line_strip) and not line_strip.startswith("On Error") and not line_strip.endswith(":"):
            print(f"[{i+1}] {line_strip}")

