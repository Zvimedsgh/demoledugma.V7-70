import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_clean.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "If isDemoMode Then" in line and "ApplyCorrectionsAndBuildReports" in "".join(lines[max(0, i-150):i]):
        for j in range(i, i+15):
            print(f"[{j}] {lines[j].strip()}")
        break
