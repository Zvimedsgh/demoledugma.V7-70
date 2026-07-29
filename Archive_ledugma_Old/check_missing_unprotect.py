with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

subs_to_fix = ['HideInstallationInstructions', 'HideInstallInst', 'HideOpInst', 'ShowResultSheets']

for sub in subs_to_fix:
    match = re.search(r'(Sub ' + sub + r'\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
    if match:
        print(f"--- {sub} ---")
        for i, line in enumerate(match.group(1).split('\n')):
            print(f"{i+1}: {line.strip()}")
