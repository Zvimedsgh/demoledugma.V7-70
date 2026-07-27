import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.168.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub AssignButtonMacros" in line or "Sub AssignButtonMacros" in line:
        for j in range(i, min(len(lines), i+100)):
            print(f"[{j+1}] {lines[j].strip()}")
            if "End Sub" in lines[j]: break
        break

