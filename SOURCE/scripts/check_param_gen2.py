import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "SetupSettingsMenu(" in line or "SetupManagementSheet(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 500: break
        if "PREMIUM_THRESHOLD" in line:
            print(f"[{i+1}] {line.strip()}")
        if "End Sub" in line and not "Setup" in line:
            break

