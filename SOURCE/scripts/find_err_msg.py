import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_20260702_1430.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "MsgBox" in line and ("A00_SetupMainSheet" in line or "Line:" in line):
        if "Line:" in line and "שגיאה" in line:
            print("Found via Hebrew check:")
            for j in range(max(0, i-10), min(len(lines), i+5)):
                print(f"[{j}] {lines[j].strip()}")
            break
        elif "Line:" in line and "MsgBox" in line:
            print("Found via MsgBox check:")
            for j in range(max(0, i-10), min(len(lines), i+5)):
                print(f"[{j}] {lines[j].strip()}")
            break
