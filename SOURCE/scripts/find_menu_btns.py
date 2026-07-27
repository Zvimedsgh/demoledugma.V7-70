import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SetupSettingsMenu()" in line:
        in_func = True
    if in_func and "End Sub" in line:
        print(f"[{i}] {line.strip()}")
        in_func = False
    
    if in_func:
        print(f"[{i}] {line.strip()}")
