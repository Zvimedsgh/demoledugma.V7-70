import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_20260702_1430.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3800, 3900):
    if "On Error" in lines[i] or "Err." in lines[i] or "wsParams" in lines[i] or "Delete" in lines[i] or "Add" in lines[i]:
        print(f"[{i}] {lines[i].strip()}")
