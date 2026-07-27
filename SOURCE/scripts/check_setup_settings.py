import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_setup = False
for i, line in enumerate(lines):
    if "Public Sub SetupSettingsSheet" in line:
        in_setup = True
    if in_setup:
        print(f"[{i+1}] {line.strip()}")
    if in_setup and "End Sub" in line:
        break

