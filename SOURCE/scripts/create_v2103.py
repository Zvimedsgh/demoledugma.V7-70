import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.102.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Fix BuildPresentation to use GetActiveAgencyName
def repl_pres_title(match):
    return match.group(0).replace('ChrW(1500) & ChrW(1489) & ChrW(1489) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495)', 'GetActiveAgencyName()')

content = re.sub(r'shp\.TextFrame\.TextRange\.Text = ChrW\(1500\).*?ChrW\(1495\)', repl_pres_title, content)

def repl_levav_name(match):
    return """Dim levavName As String
levavName = GetActiveAgencyName()"""

content = re.sub(r'Dim levavName As String\s*\n\s*levavName = ChrW\(1500\) & ChrW\(1489\) & ChrW\(1489\)', repl_levav_name, content)


# 2. Fix A00_SetupMainSheet so it doesn't show a MsgBox when it succeeds!
# We don't want a MsgBox on "Back to Home".
# It currently has:
# 4950 MsgBoxU ThisWorkbook.Worksheets(H_SET_MESSAGES()).Cells(12, 1).Value, vbInformation
# Let's remove it.
content = re.sub(r'4950 MsgBoxU ThisWorkbook\.Worksheets\(H_SET_MESSAGES\(\)\)\.Cells\(12, 1\)\.Value, vbInformation\s*\n', '', content)


# 3. Fix A00_SetupMainSheet to NOT call Application.Goto A1 at the end.
content = content.replace('4910 Application.Goto wsMain.Range("A1")', '4910 Application.Goto wsMain.Range("F3")')


# 4. Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_102"', 'Attribute VB_Name = "Goren_Claude_V2_103"')
content = content.replace('VERSION: V2.102', 'VERSION: V2.103')
content = content.replace('APP_VERSION As String = "2.102"', 'APP_VERSION As String = "2.103"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.103.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.103 created.")
