import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

found = False
for i, line in enumerate(lines):
    if "הוראות_תפעול" in line or "1514" in line: # looking for ChrW combinations
        if "BuildStructure" in "".join(lines[max(0, i-50):i]):
            pass # just a heuristic, let's just search the file for where we ADD sheets
