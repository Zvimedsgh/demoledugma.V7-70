import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_20260702_1430.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "שגיאה בהגדרת דף הבית" in line or "ERR_HANDLER:" in line:
        for j in range(i-5, i+15):
            if j < len(lines):
                print(f"[{j}] {lines[j].strip()}")
        break
