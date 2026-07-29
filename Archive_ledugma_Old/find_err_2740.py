with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Sub BuildSummarySheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for i, line in enumerate(sub_text.split('\n')):
        if '2740' in line:
            # print surrounding lines
            for j in range(max(0, i-5), min(len(sub_text.split('\n')), i+6)):
                print(f"Line {j}: {sub_text.split('\n')[j].strip()}")
