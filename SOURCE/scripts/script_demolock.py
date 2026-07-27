import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = '''' ============================================================================
' HELPER: Get Active Agency Name (Returns Demo Agency Name if in Demo Mode)'''

new_code = '''' ============================================================================
' HELPER: Check if action should be blocked in Demo Mode
' ============================================================================
Public Function CheckDemoLock() As Boolean
    Dim isDemoMode As Boolean
    Dim demoParam As String
    isDemoMode = FORCE_DEMO_MODE
    On Error Resume Next
    demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
    On Error GoTo 0
    
    If isDemoMode Then
        MsgBoxU ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497) & " " & ChrW(1488) & ChrW(1508) & ChrW(1513) & ChrW(1512) & ChrW(1497) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & ".", vbInformation, ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1491) & ChrW(1502) & ChrW(1493)
        CheckDemoLock = True
    Else
        CheckDemoLock = False
    End If
End Function

' ============================================================================
' HELPER: Get Active Agency Name (Returns Demo Agency Name if in Demo Mode)'''

if old_code not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code, new_code)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
