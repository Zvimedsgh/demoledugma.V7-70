import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.107.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

matches = re.findall(r'.{0,50}G8.{0,50}', content, re.IGNORECASE)
for m in matches:
    print(m.strip())
matches = re.findall(r'.{0,50}G9.{0,50}', content, re.IGNORECASE)
for m in matches:
    print(m.strip())

