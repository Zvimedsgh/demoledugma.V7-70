import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.140.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "If UCase(cName) Like UCase(searchText) Then" in line:
        for j in range(max(0, i-3), i+3):
            print(repr(lines[j]))

