import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.192.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_build = False
for i, line in enumerate(lines):
    if "Private Sub BuildComparisonSheet" in line:
        in_build = True
    if in_build and ("filterCol" in line or "filterVal" in line):
        print(f"[{i+1}] {line.strip()}")
    if in_build and "End Sub" in line:
        in_build = False

