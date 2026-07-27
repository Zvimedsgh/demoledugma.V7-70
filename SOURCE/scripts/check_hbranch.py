import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.192.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Function H_BRANCH() As String" in line:
        for j in range(i, i+3):
            print(f"[{j+1}] {lines[j].strip()}")
    if "Public Function H_MAINBRANCH() As String" in line:
        for j in range(i, i+3):
            print(f"[{j+1}] {lines[j].strip()}")

