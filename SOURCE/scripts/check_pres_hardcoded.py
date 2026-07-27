import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.103.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern = re.compile(r'Public Sub BuildPresentation\(\).*?End Sub', re.DOTALL)
match = pattern.search(content)
if match:
    sub_content = match.group(0)
    for i, line in enumerate(sub_content.splitlines()):
        if "ChrW(1500) & ChrW(1489) & ChrW(1489)" in line:
            print(f"[{i+1}] {line.strip()}")

