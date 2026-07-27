import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.165.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "shpDemoMsgText" in line or "E19:K20" in line:
        for j in range(max(0, i-10), min(len(lines), i+20)):
            print(f"[{j+1}] {lines[j].strip()}")
        break
