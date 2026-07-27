import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.108.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Make GetActiveAgencyName volatile
def make_volatile(match):
    return match.group(0) + "\n    Application.Volatile"
content = re.sub(r'Public Function GetActiveAgencyName\(\) As String', make_volatile, content)

# Change .Value to .Formula
content = content.replace('wsMain.Range("A1").Value = GetActiveAgencyName()', 'wsMain.Range("A1").Formula = "=GetActiveAgencyName()"')
content = content.replace('wsMainUI.Range("A1").Value = GetActiveAgencyName()', 'wsMainUI.Range("A1").Formula = "=GetActiveAgencyName()"')

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_108"', 'Attribute VB_Name = "Goren_Claude_V2_109"')
content = content.replace('VERSION: V2.108', 'VERSION: V2.109')
content = content.replace('APP_VERSION As String = "2.108"', 'APP_VERSION As String = "2.109"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.109.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.109 created.")
