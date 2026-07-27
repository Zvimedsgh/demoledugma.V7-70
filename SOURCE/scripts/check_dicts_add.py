import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.114.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "dictTellers.Add" in line or "dictAgents.Add" in line or "dictCompanies.Add" in line:
        for j in range(max(0, i-3), min(len(lines), i+4)):
            print(f"[{j+1}] {lines[j].strip()}")
        print("-" * 20)

