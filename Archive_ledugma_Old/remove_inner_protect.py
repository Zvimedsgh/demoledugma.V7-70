with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# Remove Protect from BuildComparisonSheet
text = re.sub(r'1790\s+ThisWorkbook\.Protect\s+"Z961814r"\s*\n', '\n', text, flags=re.IGNORECASE)

# Remove Protect from BuildSummarySheet
text = re.sub(r'410\s+ThisWorkbook\.Protect\s+"Z961814r"\s*\n', '\n', text, flags=re.IGNORECASE)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Removed inner Protects!")
