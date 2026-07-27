import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038_20260702_1440.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMgmt.Range(\"B2\").Value" in line or "wsMgmt.Shapes.AddShape" in line:
        if "B2" in line or "Top" in line:
            print(f"[{i}] {lines[i].strip()}")
