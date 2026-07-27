import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.103.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace A1 with F3 for Home page Goto's
content = content.replace('Application.Goto wsMain.Range("A1")', 'Application.Goto wsMain.Range("F3")')
content = content.replace('Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("A1")', 'Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F3")')

# Fix NavToIndex and NavSettings_Home to use Goto instead of Select
content = content.replace('ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F3").Select', 'Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F3")')

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_103"', 'Attribute VB_Name = "Goren_Claude_V2_104"')
content = content.replace('VERSION: V2.103', 'VERSION: V2.104')
content = content.replace('APP_VERSION As String = "2.103"', 'APP_VERSION As String = "2.104"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.104.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.104 created.")
