import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMgmt.Cells" in line or "wsMgmt.Range" in line:
        if i < 3000:
            print(f"[{i}] {lines[i].strip()}")
