import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.181.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "For Each ws In " in line or "For Each wsMain In " in line or "ws.Visible = xlSheetVeryHidden" in line:
        pass
    if "ws.Visible = xlSheetVeryHidden" in line:
        print(f"Hiding at line {i+1}:")
        for j in range(i-5, i+2):
            if j >= 0:
                print(f"[{j+1}] {lines[j].strip()}")
        print("-" * 20)

