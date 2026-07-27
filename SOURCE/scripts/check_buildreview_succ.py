import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.107.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern = re.compile(r'Public Sub BuildReview\(\).*?End Sub', re.DOTALL)
match = pattern.search(content)
if match:
    sub_content = match.group(0)
    lines = sub_content.splitlines()
    for i in range(len(lines)-40, len(lines)):
        print(f"[{i+1}] {lines[i].strip()}")

