import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.127.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "shpDemoMsgText" in line and "A00_SetupMainSheet" in "".join(lines[max(0, i-200):i]):
        print(f"[{i+1}] {lines[i].strip()}")

