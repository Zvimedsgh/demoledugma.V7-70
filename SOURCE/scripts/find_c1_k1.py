import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038_20260702_1440.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMain.Range(\"C1:K1\").Merge" in line:
        for j in range(i-2, i+10):
            print(f"[{j}] {lines[j].strip()}")
        break
