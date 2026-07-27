import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.029.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
for i, line in enumerate(lines):
    if "Public Sub BuildSummarySheet" in line:
        in_sub = True
    if in_sub and "End Sub" in line:
        in_sub = False
    
    if in_sub and "wsMain.Range" in line:
        print(f"[{i}] {line.strip()}")
