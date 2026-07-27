import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.188.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ws.Visible = xlSheetVeryHidden" in line or "ws.Visible = xlSheetHidden" in line:
        print(f"Hiding at line {i+1}:")
        for j in range(i-3, i+2):
            if j >= 0:
                print(f"[{j+1}] {lines[j].strip()}")
        print("-" * 20)

