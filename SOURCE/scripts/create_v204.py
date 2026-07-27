import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.203.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.204.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

old_block = """4600 wsParams.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = ChrW(1499) & ChrW(1503) ' YES
4610 paramLastRow = paramLastRow + 1
4620 End If"""

new_block = """4600 wsParams.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = ChrW(1499) & ChrW(1503) ' YES
4610 paramLastRow = paramLastRow + 1
4620 End If

' Add Data Validation for DEMO_MODE
Dim demoRow As Long
If foundDemo Then demoRow = pr Else demoRow = paramLastRow - 1
On Error Resume Next
wsParams.Cells(demoRow, COL_PARAM_VALUE).Validation.Delete
wsParams.Cells(demoRow, COL_PARAM_VALUE).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=ChrW(1499) & ChrW(1503) & "," & ChrW(1500) & ChrW(1488)
On Error GoTo ERR_HANDLER"""

content = content.replace(old_block, new_block)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_203"', 'Attribute VB_Name = "Goren_Claude_V2_204"')
content = content.replace('VERSION: V2.203', 'VERSION: V2.204')
content = content.replace('APP_VERSION As String = "2.203"', 'APP_VERSION As String = "2.204"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.204")
