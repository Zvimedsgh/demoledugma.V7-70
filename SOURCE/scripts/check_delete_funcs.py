import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.120.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsCleanup.Delete" in line or "wsTarget.Delete" in line:
        for j in range(max(0, i-25), i+1):
            if "Sub " in lines[j] or "Function " in lines[j]:
                print(f"In Function: {lines[j].strip()}")
                break
        print(f"[{i+1}] {line.strip()}")
        print("-" * 20)

