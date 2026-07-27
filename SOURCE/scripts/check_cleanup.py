import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.072.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMain.Range(\"A2:K20\").ClearContents" in line:
        for j in range(i-10, i+15):
            print(f"[{j+1}] {lines[j].strip()}")

