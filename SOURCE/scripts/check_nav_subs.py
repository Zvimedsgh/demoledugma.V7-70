import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.099.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

matches = re.findall(r'Public Sub Nav[A-Za-z0-9_]*\(\).*?End Sub', content, re.DOTALL)
for m in matches:
    print(m)
    print("-" * 40)

