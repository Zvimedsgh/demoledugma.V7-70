import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.137.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsSearch.Activate" in line:
        print(f"[{i+1}] {lines[i].strip()}")
    if "wsSearch.Visible" in line and "SearchClientName" not in line:
        pass # just checking

