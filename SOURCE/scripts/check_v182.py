import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.182.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "CreateInstructionSheets" in line:
        print(f"Found CreateInstructionSheets at {i+1}")
    if "Dim hideIt As Boolean" in line or "Dim hideIt2 As Boolean" in line:
        print(f"Found replacement loop at {i+1}")

