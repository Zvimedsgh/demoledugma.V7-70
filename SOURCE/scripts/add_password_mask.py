import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.117.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Insert APIs at the top
api_code = """
' --- Windows API for Password InputBox ---
#If VBA7 Then
Private Declare PtrSafe Function FindWindowExW Lib "user32" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As LongPtr, ByVal lpsz2 As LongPtr) As LongPtr
Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal Hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal Hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As Any) As LongPtr
#Else
Private Declare Function FindWindowExW Lib "user32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As Long, ByVal lpsz2 As Long) As Long
Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal Hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal Hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
#End If
Private Const EM_SETPASSWORDCHAR = &HCC
"""

if "FindWindowExW" not in content:
    content = content.replace("' --- General constants ---", api_code + "\n' --- General constants ---")

# 2. Insert TimerProc right before ToggleHiddenSheets
timer_proc = """
#If VBA7 Then
Public Sub PasswordTimerProc(ByVal Hwnd As LongPtr, ByVal uMsg As Long, ByVal nIDEvent As LongPtr, ByVal dwTimer As Long)
    Dim hwndDlg As LongPtr, hwndEdit As LongPtr
#Else
Public Sub PasswordTimerProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTimer As Long)
    Dim hwndDlg As Long, hwndEdit As Long
#End If
    On Error Resume Next
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
    End If
End Sub

"""

if "PasswordTimerProc" not in content:
    content = content.replace("Public Sub ToggleHiddenSheets()", timer_proc + "Public Sub ToggleHiddenSheets()")

# 3. Add SetTimer call before InputBox
old_input = """Dim pwd As String
pwd = InputBox(ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " " & ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & ":", ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492))"""

new_input = """Dim pwd As String
SetTimer 0, 0, 10, AddressOf PasswordTimerProc
pwd = InputBox(ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " " & ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & ":", ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492))"""

content = content.replace(old_input, new_input)


# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_117"', 'Attribute VB_Name = "Goren_Claude_V2_118"')
content = content.replace('VERSION: V2.117', 'VERSION: V2.118')
content = content.replace('APP_VERSION As String = "2.117"', 'APP_VERSION As String = "2.118"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.118.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.118 created.")
