import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if ".OnAction" in line:
        if "NavTo" in line or "Settings" in line or "Setup" in line or "Mgmt" in line or "Manage" in line:
            print(f"[{i}] {line.strip()}")
