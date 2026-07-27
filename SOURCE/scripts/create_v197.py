import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.196.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.197.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix UpdateFilterValueDropdown Validation.Add
old_code = """APPLY_VALIDATION:
' Add validation list to G10
400     wsMain.Range("rngFilterValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=valList
' Set default value and jump cursor
405     wsMain.Range("rngFilterValue").Value = selectText"""

new_code = """APPLY_VALIDATION:
' Add validation list to G10
On Error Resume Next
400     wsMain.Range("rngFilterValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=valList
On Error GoTo ERR_HANDLER
' Set default value and jump cursor
405     wsMain.Range("rngFilterValue").Value = selectText"""
content = content.replace(old_code, new_code)


# Fix UpdateFilterValueDropdown wsLists.ClearContents
old_clear = """Set wsLists = ThisWorkbook.Worksheets(listsName)
wsLists.Range("T:T").ClearContents"""

new_clear = """Set wsLists = ThisWorkbook.Worksheets(listsName)
On Error Resume Next
wsLists.Unprotect "Z961814r"
On Error GoTo ERR_HANDLER
wsLists.Range("T:T").ClearContents"""
content = content.replace(old_clear, new_clear)

# Update version strings
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_196"', 'Attribute VB_Name = "Goren_Claude_V2_197"')
content = content.replace('VERSION: V2.196', 'VERSION: V2.197')
content = content.replace('APP_VERSION As String = "2.196"', 'APP_VERSION As String = "2.197"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.197")
