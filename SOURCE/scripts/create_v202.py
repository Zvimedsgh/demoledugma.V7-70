import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.201.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.202.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix SettingsGoToSheet
old_goto = """Private Sub SettingsGoToSheet(shtName As String)
On Error Resume Next
Dim ws As Worksheet"""

new_goto = """Private Sub SettingsGoToSheet(shtName As String)
On Error Resume Next
ThisWorkbook.Unprotect "Z961814r"
Dim ws As Worksheet"""

content = content.replace(old_goto, new_goto)

# Fix UpdateFilterValueDropdown Demo Mode logic
old_demo = """If Not wsDataUI Is Nothing Then
Dim dCol As Long
If filterType = H_COMPANY() Then
dCol = BASE_COL_COMPANY
ElseIf filterType = H_TELLER() Then
dCol = BASE_COL_TELLER
ElseIf filterType = H_AGENT() Then
dCol = BASE_COL_AGENTNAME
ElseIf filterType = H_BRANCH() Then
dCol = BASE_COL_BRANCHNAME
ElseIf filterType = H_BRANCH() & ChrW(32) & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494) Then
dCol = BASE_COL_MAINBRANCH
Else
GoTo CLEAN_EXIT
End If"""

new_demo = """If Not wsDataUI Is Nothing Then
InitRawColumns
Dim dCol As Long
If filterType = H_COMPANY() Then
dCol = RAW_COMPANY
ElseIf filterType = H_TELLER() Then
dCol = RAW_TELLERNAME
ElseIf filterType = H_AGENT() Then
dCol = RAW_AGENTNAME
ElseIf filterType = H_BRANCH() Then
dCol = RAW_BRANCHNAME
ElseIf filterType = H_BRANCH() & ChrW(32) & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494) Then
dCol = RAW_BRANCHNAME
Else
GoTo CLEAN_EXIT
End If"""

content = content.replace(old_demo, new_demo)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_201"', 'Attribute VB_Name = "Goren_Claude_V2_202"')
content = content.replace('VERSION: V2.201', 'VERSION: V2.202')
content = content.replace('APP_VERSION As String = "2.201"', 'APP_VERSION As String = "2.202"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.202")
