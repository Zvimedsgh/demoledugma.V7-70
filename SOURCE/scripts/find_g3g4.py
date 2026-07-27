import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMain.Range(\"G3\").Value" in line or "wsMain.Range(\"G4\").Value" in line:
        print(f"[{i}] {line.strip()}")
