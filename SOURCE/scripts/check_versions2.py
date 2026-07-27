import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.137.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Attribute VB_Name" in line or "VERSION" in line or "APP_VERSION As String =" in line:
        print(f"[{i+1}] {lines[i].strip()}")

