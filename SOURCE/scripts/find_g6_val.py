import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMain.Range(\"G6\").MergeArea.Validation.Add" in line or "wsMain.Range(\"G9\").MergeArea.Validation.Add" in line:
        for j in range(i-2, i+3):
            print(f"[{j}] {lines[j].strip()}")
