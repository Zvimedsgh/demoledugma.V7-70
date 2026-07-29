import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.04.bas', 'r', encoding='utf-8') as f:
    content = f.read()

old_b7_code = """    ' Add Yes/No Dropdown in Settings B7
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
    On Error GoTo ERR_HANDLER"""

new_b7_code = """    ' Add Yes/No Dropdown dynamically to DEMO_MODE value cell (Column C)
    On Error Resume Next
    Dim wsParamsSetup As Worksheet
    Set wsParamsSetup = ThisWorkbook.Worksheets(H_SET_PARAMS())
    If Not wsParamsSetup Is Nothing Then
        wsParamsSetup.Unprotect Password:="1234"
        Dim rParam As Long
        For rParam = 2 To 100
            If UCase$(Trim$(CStr(wsParamsSetup.Cells(rParam, 2).Value2))) = "DEMO_MODE" Then
                wsParamsSetup.Cells(rParam, 3).Locked = False
                With wsParamsSetup.Cells(rParam, 3).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=ChrW(1499) & ChrW(1503) & "," & ChrW(1500) & ChrW(1488)
                    .IgnoreBlank = True
                    .InCellDropdown = True
                End With
                Exit For
            End If
        Next rParam
        wsParamsSetup.Protect Password:="1234"
    End If
    On Error GoTo ERR_HANDLER"""

content = content.replace(old_b7_code, new_b7_code)

content = content.replace('APP_VERSION As String = "3.04"', 'APP_VERSION As String = "3.05"')
content = content.replace('VERSION: V3.04', 'VERSION: V3.05')
content = content.replace('Error in V3.04!', 'Error in V3.05!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_04"', 'Attribute VB_Name = "Goren_Claude_V3_05"')

changelog = """' CHANGES IN 3.05:
'   - BUGFIX: Moved DEMO_MODE Yes/No validation from B7 (the label) to C7 (the value), and unlocked the cell so it can be clicked while the sheet is protected.
"""
content = content.replace("' CHANGES IN 3.04:", changelog + "' CHANGES IN 3.04:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.05.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.05')
