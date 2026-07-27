import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.185.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.186.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update ToggleHiddenSheets password
old_pwd = """    If anyOtherVisible Then
        ' Hiding sheets - no password required
        HideWorkSheets
    Else
        ' Showing sheets - temporarily disabled password for debugging
        ShowHiddenSheets
    End If"""

new_pwd = """    If anyOtherVisible Then
        ' Hiding sheets - no password required
        HideWorkSheets
    Else
        ' Showing sheets - check password
        Dim psw As String
        #If VBA7 Then
            SetTimer 0, 0, 10, AddressOf PasswordTimerProc
        #Else
            SetTimer 0, 0, 10, AddressOf PasswordTimerProc
        #End If
        psw = InputBox(ChrW(1492) & ChrW(1499) & ChrW(1504) & ChrW(1505) & " " & ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & " " & ChrW(1499) & ChrW(1491) & ChrW(1497) & " " & ChrW(1500) & ChrW(1492) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1493) & ChrW(1505) & ChrW(1514) & ChrW(1512) & ChrW(1497) & ChrW(1501) & ":", ChrW(1504) & ChrW(1491) & ChrW(1512) & ChrW(1513) & ChrW(1514) & " " & ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492))
        If psw = ADMIN_PASSWORD() Then
            ShowHiddenSheets
        Else
            MsgBoxU ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & " " & ChrW(1513) & ChrW(1490) & ChrW(1493) & ChrW(1497) & ChrW(1492), vbCritical
        End If
    End If"""

content = content.replace(old_pwd, new_pwd)

# 2. Update CheckUserPermissions activation
old_act = """    ' Navigate to home page
    ThisWorkbook.Worksheets(homeSheetName).Activate"""

new_act = """    ' Navigate to home page
    ThisWorkbook.Worksheets(homeSheetName).Activate
    
    ' But if instruction sheets are enabled, jump to them instead
    Dim sInstall2 As String
    sInstall2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
    If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA1").Value) <> "YES" Then
        On Error Resume Next
        ThisWorkbook.Worksheets(sInstall2).Activate
        On Error GoTo 0
    End If"""

content = content.replace(old_act, new_act)


content = content.replace('Attribute VB_Name = "Goren_Claude_V2_185"', 'Attribute VB_Name = "Goren_Claude_V2_186"')
content = content.replace('VERSION: V2.185', 'VERSION: V2.186')
content = content.replace('APP_VERSION As String = "2.185"', 'APP_VERSION As String = "2.186"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.186 correctly!")
