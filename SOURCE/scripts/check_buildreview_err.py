import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.107.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern = re.compile(r'Public Sub BuildReview\(\).*?End Sub', re.DOTALL)
match = pattern.search(content)
if match:
    sub_content = match.group(0)
    for i, line in enumerate(sub_content.splitlines()):
        if "ERR_HANDLER:" in line or "CLEAN_EXIT:" in line:
            for j in range(max(0, i-2), len(sub_content.splitlines())):
                print(f"[{j+1}] {sub_content.splitlines()[j].strip()}")
            break

