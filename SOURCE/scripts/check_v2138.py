import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

found_activate = False
found_dir = 0
for i, line in enumerate(lines):
    if "wsSearch.Visible = xlSheetVisible" in line:
        found_activate = True
    if "TextDirection = 2" in line:
        found_dir += 1

print(f"Activate fix: {found_activate}")
print(f"Text direction fix: {found_dir}")

