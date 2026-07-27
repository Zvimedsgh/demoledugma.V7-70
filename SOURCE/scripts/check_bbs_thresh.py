import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "If Abs(premVal) > threshold And Not dictCorrections.Exists(CStr(r)) Then GoTo NextSrcRow" in line:
        print(f"[{i+1}] {line.strip()}")

