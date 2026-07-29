with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Public Sub ApplyCorrectionsAndBuildReports\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    lines = sub_text.split('\n')
    for i in range(125, 135):
        print(f"Line {i+1}: {lines[i].strip()}")
