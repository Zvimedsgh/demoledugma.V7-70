import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.158.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Dim askMsg As String" in line:
        for j in range(i, i+50):
            print(f"[{j+1}] {lines[j].strip()}")
        break

