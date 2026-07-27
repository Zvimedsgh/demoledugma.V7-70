import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        in_func = True
    if in_func and "End Sub" in line:
        in_func = False
    
    if in_func and "SetupSettingsMenu" in line:
        print(f"[{i}] {line.strip()}")
