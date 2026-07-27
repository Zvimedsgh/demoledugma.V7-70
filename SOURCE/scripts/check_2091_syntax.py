import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.091.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "End Sub" in line or "End Function" in line or "End Property" in line:
        # Check next few lines
        pass
    if "Private Sub LogDebug" in line:
        print(f"LogDebug found at line {i+1}")
        for j in range(max(0, i-5), min(len(lines), i+20)):
            print(f"[{j+1}] {lines[j].strip()}")

