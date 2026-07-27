import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.135.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "shpDemoMsg.TextFrame2.TextRange.ParagraphFormat.Alignment" in line:
        print(f"[{i+1}] {lines[i].strip()}")

