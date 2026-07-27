import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.198.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "On Error GoTo ERR_HANDLER" in line and i > 4470 and i < 4500:
        for j in range(i-5, i+5):
            print(f"[{j+1}] {lines[j].strip()}")

