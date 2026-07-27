import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = '''    If IsEmpty(wsMain.Range("G3").Value) Then wsMain.Range("G3").Value = 2025
    If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2026'''

new_code = '''    If IsEmpty(wsMain.Range("G3").Value) Then wsMain.Range("G3").Value = 2025
    If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2026

    Dim isDemoMode As Boolean
    Dim demoParam As String
    isDemoMode = FORCE_DEMO_MODE
    On Error Resume Next
    demoParam = UCase$(Trim$(GetStringParameter(wsMgmt, "DEMO_MODE")))
    If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
    On Error GoTo ERR_HANDLER

    ' Lock years if demo mode
    If isDemoMode Then
        With wsMain.Range("G3:G4").Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
            .ErrorMessage = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497) & " " & ChrW(1488) & ChrW(1508) & ChrW(1513) & ChrW(1512) & ChrW(1497) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & "."
            .ShowError = True
        End With
        wsMain.Range("G3:G4").Interior.Color = RGB(230, 230, 230) ' Gray out visually
    Else
        wsMain.Range("G3:G4").Validation.Delete
        wsMain.Range("G3:G4").Interior.Color = xlNone
    End If'''

if old_code not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code, new_code)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
