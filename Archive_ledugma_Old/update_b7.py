import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.02.bas', 'r', encoding='utf-8') as f:
    content = f.read()

target_text = "4940 ApplyDemoLockOnOpen\n4960 Exit Sub"
replacement_text = """4940 ApplyDemoLockOnOpen

    ' Add Yes/No Dropdown in Settings B7
    On Error Resume Next
    Dim wsParamsSetup As Worksheet
    Set wsParamsSetup = ThisWorkbook.Worksheets(H_SET_PARAMS())
    If Not wsParamsSetup Is Nothing Then
        wsParamsSetup.Unprotect Password:="1234"
        With wsParamsSetup.Range("B7").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=ChrW(1499) & ChrW(1503) & "," & ChrW(1500) & ChrW(1488)
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        wsParamsSetup.Protect Password:="1234"
    End If
    On Error GoTo ERR_HANDLER

4960 Exit Sub"""

content = content.replace(target_text, replacement_text)

content = content.replace('APP_VERSION As String = "3.02"', 'APP_VERSION As String = "3.03"')
content = content.replace('VERSION: V3.02', 'VERSION: V3.03')
content = content.replace('Error in V3.02!', 'Error in V3.03!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_02"', 'Attribute VB_Name = "Goren_Claude_V3_03"')

changelog = """' CHANGES IN 3.03:
'   - UI: Fixed bug where B7 dropdown was not injected because of a mismatch in python script string replacement.
"""
content = content.replace("' CHANGES IN 3.02:", changelog + "' CHANGES IN 3.02:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.03.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.03')
