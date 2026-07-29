with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Public Sub ApplyCorrectionsAndBuildReports\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    lines = sub_text.split('\n')
    for i, line in enumerate(lines):
        if 'BuildSummarySheet' in line:
            for j in range(i, len(lines)):
                print(f"Line {j+1}: {lines[j].strip()}")
            break
