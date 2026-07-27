import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "PREMIUM_THRESHOLD" in line:
        print(f"[{i+1}] {line.strip()}")
    if "BuildParamSheet" in line:
        for j in range(i, i+30):
            if "PREMIUM_THRESHOLD" in lines[j] or "200000" in lines[j] or "THRESHOLD" in lines[j]:
                print(f"[{j+1}] {lines[j].strip()}")

