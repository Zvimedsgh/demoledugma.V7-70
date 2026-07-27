import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.118.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_timer = """    On Error Resume Next
    KillTimer 0, nIDEvent
    
    Dim title As String
    title = ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492)
    
    ' Find the InputBox window by title
    hwndDlg = FindWindowW(0, StrPtr(title))
    If hwndDlg <> 0 Then
        hwndEdit = FindWindowExW(hwndDlg, 0, StrPtr("Edit"), 0)
        If hwndEdit <> 0 Then
            SendMessage hwndEdit, EM_SETPASSWORDCHAR, Asc("*"), ByVal 0&
        End If
    End If"""

new_timer = """    On Error Resume Next
    Static attempt As Long
    attempt = attempt + 1
    
    Dim title As String
    title = ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492)
    
    ' Find the InputBox window by title
    hwndDlg = FindWindowW(0, StrPtr(title))
    If hwndDlg <> 0 Then
        hwndEdit = FindWindowExW(hwndDlg, 0, StrPtr("Edit"), 0)
        If hwndEdit <> 0 Then
            SendMessage hwndEdit, EM_SETPASSWORDCHAR, Asc("*"), ByVal 0&
        End If
        KillTimer 0, nIDEvent
        attempt = 0
    ElseIf attempt > 50 Then
        KillTimer 0, nIDEvent
        attempt = 0
    End If"""

content = content.replace(old_timer, new_timer)

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_118"', 'Attribute VB_Name = "Goren_Claude_V2_119"')
content = content.replace('VERSION: V2.118', 'VERSION: V2.119')
content = content.replace('APP_VERSION As String = "2.118"', 'APP_VERSION As String = "2.119"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.119 created.")
