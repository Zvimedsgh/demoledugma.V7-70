import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.040_20260702_1507.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMain.Cells.Clear" in line or "wsMain.DrawingObjects.Delete" in line:
        print(f"[{i}] {lines[i].strip()}")
