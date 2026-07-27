import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.101.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "For Each ws In " in line or "ws.Visible = xlSheetVisible" in line:
        for j in range(max(0, i-2), min(len(lines), i+3)):
            if "For Each" in lines[j] or "Visible" in lines[j]:
                print(f"[{j+1}] {lines[j].strip()}")

