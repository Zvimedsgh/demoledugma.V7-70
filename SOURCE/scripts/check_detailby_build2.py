import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.192.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_build = False
for i, line in enumerate(lines):
    if "Public Sub BuildPresentation" in line:
        in_build = True
    if in_build and ("rngFilterType" in line or "rngFilterValue" in line or "filterValue" in line or "filterType" in line):
        print(f"[{i+1}] {line.strip()}")
    if in_build and "End Sub" in line:
        in_build = False

