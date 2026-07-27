import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = '''' ============================================================================
' HELPER: Get numeric parameter from NIHUL'''
new_code = '''' ============================================================================
' HELPER: Get Active Agency Name (Returns Demo Agency Name if in Demo Mode)
' ============================================================================
Public Function GetActiveAgencyName() As String
    Dim isDemoMode As Boolean
    Dim demoParam As String
    
    isDemoMode = FORCE_DEMO_MODE
    On Error Resume Next
    demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
    
    If isDemoMode Then
        GetActiveAgencyName = Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), PARAM_DEMO_AGENCY_NAME))
    Else
        GetActiveAgencyName = Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), PARAM_AGENCY_NAME))
    End If
    On Error GoTo 0
    
    ' Fallback to default if empty
    If GetActiveAgencyName = "" Then
        GetActiveAgencyName = ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489) ' "Levav System" default
    End If
End Function

' ============================================================================
' HELPER: Get numeric parameter from NIHUL'''

if old_code not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code, new_code)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
