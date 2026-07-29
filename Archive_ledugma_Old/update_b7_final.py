import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.05.bas', 'r', encoding='utf-8') as f:
    content = f.read()

old_code = """    ' Add Yes/No Dropdown dynamically to DEMO_MODE value cell (Column C)
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

new_code = """    ' Add Yes/No Dropdown dynamically to DEMO_MODE value cell (Column B)
    On Error Resume Next
    Dim wsParamsSetup As Worksheet
    Set wsParamsSetup = ThisWorkbook.Worksheets(H_SET_PARAMS())
    If Not wsParamsSetup Is Nothing Then
        wsParamsSetup.Unprotect Password:="1234"
        Dim rParam As Long
        For rParam = 1 To 100
            If UCase$(Trim$(CStr(wsParamsSetup.Cells(rParam, 1).Value2))) = "DEMO_MODE" Then
                wsParamsSetup.Cells(rParam, 2).Locked = False
                With wsParamsSetup.Cells(rParam, 2).Validation
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

content = content.replace(old_code, new_code)

content = content.replace('APP_VERSION As String = "3.05"', 'APP_VERSION As String = "3.06"')
content = content.replace('VERSION: V3.05', 'VERSION: V3.06')
content = content.replace('Error in V3.05!', 'Error in V3.06!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_05"', 'Attribute VB_Name = "Goren_Claude_V3_06"')

changelog = """' CHANGES IN 3.06:
'   - BUGFIX: Reverted the DEMO_MODE dropdown injection back to Column B (value) instead of Column C. The user was completely right!
"""
content = content.replace("' CHANGES IN 3.05:", changelog + "' CHANGES IN 3.05:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.06.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.06')
