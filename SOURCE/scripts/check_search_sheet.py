import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.140.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SetupSearchSheet(" in line or "Private Sub SetupSearchSheet" in line or "Sub OpenSearchScreen" in line:
        in_func = True
    if in_func:
        if i - lines.index("Public Sub OpenSearchScreen()\n") > 300: break
        if "Set wsSearch =" in line or "wsSearch.Name" in line or "Worksheet_Change" in line:
            print(f"[{i+1}] {lines[i].strip()}")
        if "End Sub" in line:
            in_func = False

