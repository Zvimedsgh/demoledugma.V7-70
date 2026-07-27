import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.092.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "FetchBOIMonthlyAvg" in line:
        print(f"[{i+1}] {line.strip()}")
    if "A00_SetupMainSheet" in line:
        print(f"[{i+1}] {line.strip()}")
    if "LogDebug" in line:
        pass

