import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.185.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Password" in line or "סיסמה" in line:
        if i > 6150 and i < 6200:
            print(f"[{i+1}] {lines[i].strip()}")

