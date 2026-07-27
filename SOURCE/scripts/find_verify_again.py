import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.052_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Function IsProtectedSheet" in line:
        for j in range(i, i+25):
            print(f"[{j}] {lines[j].strip()}")
        break

for i, line in enumerate(lines):
    if "wsCleanup.Delete" in line:
        for j in range(i-10, i+5):
            print(f"[{j}] {lines[j].strip()}")
        break
