import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.102.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub BuildPresentation()" in line:
        in_func = True
    elif in_func and "End Sub" in line:
        break
    elif in_func and ("Title" in line or "TextRange" in line):
        if i % 5 == 0:
            print(f"[{i+1}] {line.strip()}")

