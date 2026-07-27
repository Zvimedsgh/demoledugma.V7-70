import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.109.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

matches = re.finditer(r'(For pg = 1 To ppPres.Slides.Count\s*\n\s*\d*\s*)(.*?)\n(\s*\d*\s*Next pg)', content, re.IGNORECASE)
for m in matches:
    print(m.group(0))

