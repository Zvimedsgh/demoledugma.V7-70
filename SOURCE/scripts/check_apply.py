import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.208.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub ApplyCorrectionsAndBuildReports" in line:
        for j in range(i, i+300):
            if "ERR_HANDLER:" in lines[j] or "Err.Description" in lines[j]:
                print(f"[{j+1}] {lines[j].strip()}")
        break
