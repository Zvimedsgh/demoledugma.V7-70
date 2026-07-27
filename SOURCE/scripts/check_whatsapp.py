import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.174.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "WhatsApp" in line or "shpDemoMsgText" in line or "AddTextbox" in line:
        print(f"[{i+1}] {line.strip()}")

