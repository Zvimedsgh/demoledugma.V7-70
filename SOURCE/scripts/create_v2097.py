import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.096.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.097.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_096"', 'Attribute VB_Name = "Goren_Claude_V2_097"')
content = content.replace("' VERSION: V2.096", "' VERSION: V2.097")
content = content.replace('Private Const APP_VERSION As String = "2.096"', 'Private Const APP_VERSION As String = "2.097"')

# Completely neuter FetchBOIMonthlyAvg
fetch_new = """Private Function FetchBOIMonthlyAvg(ByVal currencyKey As String, Optional ByVal monthStr As String = "") As Double
    FetchBOIMonthlyAvg = 0
End Function"""

pattern = re.compile(r'Private Function FetchBOIMonthlyAvg\(.*?End Function', re.DOTALL)
content = pattern.sub(fetch_new, content)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.097 created.")
