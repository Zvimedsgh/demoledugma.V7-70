import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.141.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_search_start = """Public Sub SearchClientName()
On Error GoTo ERR_HANDLER"""

new_search_start = """Public Sub SearchClientName()
On Error GoTo ERR_HANDLER
Application.EnableEvents = False ' Prevent Worksheet_Change from firing during setup"""

target_search_end = """wsSearch.Visible = xlSheetVisible
wsSearch.Activate
Application.Goto wsSearch.Cells(2, 1)
Exit Sub
ERR_HANDLER:
Application.EnableEvents = True"""

new_search_end = """wsSearch.Visible = xlSheetVisible
wsSearch.Activate
Application.Goto wsSearch.Cells(2, 1)
Application.EnableEvents = True
Exit Sub
ERR_HANDLER:
Application.EnableEvents = True"""

if target_search_start in content:
    content = content.replace(target_search_start, new_search_start)
    content = content.replace(target_search_end, new_search_end)
    print("Fixed SearchClientName events!")
else:
    print("Could not find SearchClientName start!")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_141"', 'Attribute VB_Name = "Goren_Claude_V2_142"')
content = content.replace('VERSION: V2.141', 'VERSION: V2.142')
content = content.replace('APP_VERSION As String = "2.141"', 'APP_VERSION As String = "2.142"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.142.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.142 created.")
