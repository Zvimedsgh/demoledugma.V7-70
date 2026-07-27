import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.027.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

import re
matches = re.findall(r'Dim isDemoMode As Boolean', content)
print(f"Total Dim isDemoMode As Boolean: {len(matches)}")
