import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.167.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

found = False
for i, line in enumerate(lines):
    if "shpDemoMsgText" in line:
        found = True
        print(f"[{i+1}] {line.strip()}")

if not found:
    print("Not found.")
