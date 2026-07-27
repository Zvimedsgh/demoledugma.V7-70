import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.145.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Agency" in line or "AGENCY" in line or "AgencyName" in line or "AGENCY_NAME" in line or "Name" in line and "PARAM" in line:
        print(f"[{i+1}] {line.strip()}")

