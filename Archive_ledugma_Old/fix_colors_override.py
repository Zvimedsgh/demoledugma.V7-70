with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
pattern = r'(ThisWorkbook\.Worksheets\(MATACH_SHEET_NAME\(\)\)\.Tab\.Color = RGB\(0, 80, 120\))'
replacement = r'\1\nOn Error Resume Next\nThisWorkbook.Worksheets(ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)).Tab.Color = RGB(255, 153, 153)'

text = re.sub(pattern, replacement, text)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated SetAllTabColors successfully.")
