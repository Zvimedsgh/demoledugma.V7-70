with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Sub BuildComparisonSheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    if 'Protect' in sub_text:
        print("Protect is STILL in BuildComparisonSheet!")
        for i, line in enumerate(sub_text.split('\n')):
            if 'Protect' in line:
                print(f"Line {i+1}: {line.strip()}")
    else:
        print("Protect was successfully removed from BuildComparisonSheet.")
