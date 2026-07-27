import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.198.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix UpdateFilterValueDropdown error handling around wsLists modification
old_block = """On Error Resume Next
wsLists.Unprotect "Z961814r"
On Error GoTo ERR_HANDLER
wsLists.Range("T:T").ClearContents

Dim kUI As Variant
rUI = 1
For Each kUI In collUI.Keys
rUI = rUI + 1
wsLists.Cells(rUI, 20).Value = kUI
Next kUI

ThisWorkbook.Names.Add "lst_temp_filter", wsLists.Range(wsLists.Cells(2, 20), wsLists.Cells(rUI, 20))"""

new_block = """On Error Resume Next
wsLists.Unprotect "Z961814r"
wsLists.Range("T:T").ClearContents

Dim kUI As Variant
rUI = 1
For Each kUI In collUI.Keys
rUI = rUI + 1
wsLists.Cells(rUI, 20).Value = kUI
Next kUI

ThisWorkbook.Names.Add "lst_temp_filter", wsLists.Range(wsLists.Cells(2, 20), wsLists.Cells(rUI, 20))
On Error GoTo ERR_HANDLER"""

content = content.replace(old_block, new_block)

# Fix early filterValue clearing
old_clear = """' Clear rngFilterValue
50      wsMain.Range("rngFilterValue").Value = ""
60      On Error Resume Next
70      wsMain.Range("rngFilterValue").Validation.Delete
80      On Error GoTo ERR_HANDLER"""

new_clear = """' Clear rngFilterValue
60      On Error Resume Next
50      wsMain.Range("rngFilterValue").Value = ""
70      wsMain.Range("rngFilterValue").Validation.Delete
80      On Error GoTo ERR_HANDLER"""

content = content.replace(old_clear, new_clear)

# Update version strings
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_198"', 'Attribute VB_Name = "Goren_Claude_V2_199"')
content = content.replace('VERSION: V2.198', 'VERSION: V2.199')
content = content.replace('APP_VERSION As String = "2.198"', 'APP_VERSION As String = "2.199"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.199")
