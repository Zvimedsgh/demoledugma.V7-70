import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A1_RunDataValidation" in line or "Sub RunDataValidation" in line or "Sub A1_" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 100: break
        if "BuildReview" in line or "yearVal" in line:
            print(f"[{i+1}] {line.strip()}")
        if "End Sub" in line:
            break

