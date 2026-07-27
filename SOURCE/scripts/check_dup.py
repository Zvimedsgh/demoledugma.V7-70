import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.026.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
count = 0
for i, line in enumerate(lines):
    if "Public Sub ApplyCorrectionsAndBuildReports()" in line:
        in_sub = True
    if in_sub and "End Sub" in line:
        in_sub = False
    
    if in_sub and "Dim isDemoMode As Boolean" in line:
        count += 1
        print(f"Found at line {i}: {line.strip()}")

print(f"Total in ApplyCorrectionsAndBuildReports: {count}")
