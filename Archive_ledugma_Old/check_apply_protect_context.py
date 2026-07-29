with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Public Sub ApplyCorrectionsAndBuildReports\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for i, line in enumerate(sub_text.split('\n')):
        if 'ThisWorkbook.Protect' in line:
            print(f"--- Line {i+1} ---")
            for j in range(max(0, i-2), i+1):
                print(f"  {sub_text.splitlines()[j].strip()}")
