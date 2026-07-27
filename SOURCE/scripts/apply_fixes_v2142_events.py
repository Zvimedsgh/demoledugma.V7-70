import sys, re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.141.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern_start = re.compile(r"(Public Sub SearchClientName\(\)\s*On Error GoTo ERR_HANDLER\s*)")
if pattern_start.search(content):
    content = pattern_start.sub(r"\1Application.EnableEvents = False\n    ", content)
    print("Found and replaced start of SearchClientName")
else:
    print("Failed to replace start")

pattern_end = re.compile(r"(wsSearch\.Cells\(2, 1\)\s*Exit Sub\s*ERR_HANDLER:\s*Application\.EnableEvents = True)")
if pattern_end.search(content):
    content = pattern_end.sub(r"wsSearch.Cells(2, 1)\n    Application.EnableEvents = True\n    Exit Sub\nERR_HANDLER:\n    Application.EnableEvents = True", content)
    print("Found and replaced end of SearchClientName")
else:
    print("Failed to replace end")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_141"', 'Attribute VB_Name = "Goren_Claude_V2_142"')
content = content.replace('VERSION: V2.141', 'VERSION: V2.142')
content = content.replace('APP_VERSION As String = "2.141"', 'APP_VERSION As String = "2.142"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.142.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.142 updated.")
