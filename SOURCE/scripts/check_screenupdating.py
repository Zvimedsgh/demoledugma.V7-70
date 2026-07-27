import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.137.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = ""
for i, line in enumerate(lines):
    if "Public Sub" in line or "Private Sub" in line:
        in_func = line.strip()
    if "Application.ScreenUpdating = False" in line:
        print(f"Found ScreenUpdating=False in {in_func}")

