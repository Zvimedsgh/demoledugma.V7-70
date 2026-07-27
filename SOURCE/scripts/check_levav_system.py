import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.104.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Let's just search for Levav System
matches = re.findall(r'.{0,50}Levav System.{0,50}', content, re.IGNORECASE)
for m in matches:
    print(m.strip())

