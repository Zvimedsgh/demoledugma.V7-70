import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "20" in line and ("=" in line or "," in line) and len(line.strip()) < 100:
        if "RGB" not in line and "MsgBox" not in line and "Wait" not in line:
            print(f"[{i}] {line.strip()}")
