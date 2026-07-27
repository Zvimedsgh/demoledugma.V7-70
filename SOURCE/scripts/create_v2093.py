import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.092.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.093.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_092"', 'Attribute VB_Name = "Goren_Claude_V2_093"')
content = content.replace("' VERSION: V2.092", "' VERSION: V2.093")
content = content.replace('Private Const APP_VERSION As String = "2.092"', 'Private Const APP_VERSION As String = "2.093"')

# Completely neuter FetchBOIMonthlyAvg
pattern = re.compile(r'(Private Function FetchBOIMonthlyAvg\(.*?\).*?As Double)\r?\n', re.IGNORECASE)
content = pattern.sub(r'\1\n    FetchBOIMonthlyAvg = 0\n    Exit Function\n', content)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.093 created.")
