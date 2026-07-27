import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.160.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "shpDemoMsgText" in line:
        # found it, now trace back to Sub
        for j in range(i, -1, -1):
            if "Public Sub " in lines[j] or "Private Sub " in lines[j] or "Sub " in lines[j]:
                print(f"[{j+1}] {lines[j].strip()}")
                break
        break

