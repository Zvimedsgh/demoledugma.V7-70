import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.187.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Function " in line and "Password" in line:
        print(f"[{i+1}] {line.strip()}")
    if "ADMIN_PASSWORD" in line:
        print(f"[{i+1}] {line.strip()}")

