import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.185.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "InputBox" in line and "Password" not in line:
        pass
    if "InputBox" in line:
        print(f"[{i+1}] {line.strip()}")

