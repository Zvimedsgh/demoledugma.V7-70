with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match1 = re.search(r'(Sub BuildComparisonSheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match1:
    sub_text1 = match1.group(1)
    for line in sub_text1.split('\n'):
        if 'Worksheets.Add' in line:
            print(f"BuildComparisonSheet: {line.strip()}")

match2 = re.search(r'(Sub BuildSummarySheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match2:
    sub_text2 = match2.group(1)
    for line in sub_text2.split('\n'):
        if 'Worksheets.Add' in line:
            print(f"BuildSummarySheet: {line.strip()}")
