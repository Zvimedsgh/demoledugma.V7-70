import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "בדיקת נתונים" in line or ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1514) in line:
        pass
    # I'll just check what the button is called in SetupMainSheet
    if "btnStep1.OnAction" in line or "1 -" in line:
        print(f"[{i+1}] {line.strip()}")

