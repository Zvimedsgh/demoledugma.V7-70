import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_fixed.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Private Sub SettingsGoToSheet" in line:
        for j in range(i, i+15):
            print(f"[{j}] {lines[j].strip()}")
        break
