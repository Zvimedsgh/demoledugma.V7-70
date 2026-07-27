import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.181.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Visible = xlSheetVisible" in line or "Visible = xlSheetHidden" in line or "Visible = xlSheetVeryHidden" in line:
        pass
    if "Sub ResetHomeDefaults" in line:
        print(f"ResetHomeDefaults at {i}")
    if "Function CONTROL_SHEET_NAME" in line:
        print(f"CONTROL_SHEET_NAME at {i}")

