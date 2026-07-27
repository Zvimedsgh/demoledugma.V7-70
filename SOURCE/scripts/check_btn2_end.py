import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.151.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub ApplyCorrectionsAndBuildReports(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 2000: break
        if "LogDebug \"ApplyCorrectionsAndBuildReports ENDED SUCCESSFULLY\"" in line or "End Sub" in line:
            for j in range(max(0, i-5), i+2):
                print(f"[{j+1}] {lines[j].strip()}")
            break

