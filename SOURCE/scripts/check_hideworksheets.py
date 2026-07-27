import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.213.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub HideWorkSheets" in line:
        for j in range(i, i+30):
            try:
                print(lines[j].rstrip())
            except:
                pass
        break
