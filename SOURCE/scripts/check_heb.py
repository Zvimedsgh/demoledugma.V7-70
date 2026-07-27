import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.181.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

import re
hebrew_chars = re.findall(r'[\u0590-\u05FF]', content)
print(f"Number of Hebrew characters in V2.181: {len(hebrew_chars)}")
if len(hebrew_chars) > 0:
    print("Example snippet with Hebrew:")
    # print context of first hebrew char
    idx = content.find(hebrew_chars[0])
    print(content[max(0, idx-20):min(len(content), idx+50)])

