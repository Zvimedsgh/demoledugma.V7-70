import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.100.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SendForReview()" in line:
        in_func = True
        print(f"[{i+1}] {line.strip()}")
    elif in_func and "End Sub" in line:
        for j in range(max(0, i-20), i+1):
            print(f"[{j+1}] {lines[j].strip()}")
        break
    elif in_func and i < 5500:
        print(f"[{i+1}] {line.strip()}")

