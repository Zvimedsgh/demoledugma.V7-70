import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.141.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SearchClientName()" in line:
        for j in range(i, min(len(lines), i+10)):
            print(repr(lines[j]))
        break

