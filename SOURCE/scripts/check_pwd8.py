import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.185.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Function " in line and "Input" in line:
        print(f"[{i+1}] {line.strip()}")
    if "GetAdminPassword" in line:
        print(f"[{i+1}] {line.strip()}")
    if "SetTimer" in line:
        print(f"[{i+1}] {line.strip()}")

