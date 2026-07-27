import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.201.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "isDemoMode =" in line and "FORCE_DEMO_MODE" in line:
        for j in range(i, i+60):
            try:
                print(f"[{j+1}] {lines[j].strip()}")
            except:
                print(f"[{j+1}] [Unprintable]")
        break
