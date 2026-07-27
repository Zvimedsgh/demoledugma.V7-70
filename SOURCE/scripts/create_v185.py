import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.184.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.185.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = """For Each ws In ThisWorkbook.Worksheets
Select Case ws.Name
Case CONTROL_SHEET_NAME(), MANAGEMENT_SHEET_NAME(), SHEET_COMPANIES(), _
SHEET_AGENTS(), SHEET_BRANCH(), SHEET_MAINBRANCH(), _
SHEET_TELLERS(), SHEET_MONTHS(), SHEET_SUMMARY()
' These are always visible, so ignore them
Case Else
If ws.Visible = xlSheetVisible Then
anyOtherVisible = True
Exit For
End If
End Select
Next ws"""

new_code = """For Each ws In ThisWorkbook.Worksheets
    Dim ignoreSheet As Boolean
    ignoreSheet = False
    
    If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ignoreSheet = True ' הוראות_התקנה
    If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ignoreSheet = True ' הוראות_תפעול
    
    If Not ignoreSheet Then
        Select Case ws.Name
        Case CONTROL_SHEET_NAME(), MANAGEMENT_SHEET_NAME(), SHEET_COMPANIES(), _
        SHEET_AGENTS(), SHEET_BRANCH(), SHEET_MAINBRANCH(), _
        SHEET_TELLERS(), SHEET_MONTHS(), SHEET_SUMMARY()
        ' These are always visible, so ignore them
        Case Else
        If ws.Visible = xlSheetVisible Then
        anyOtherVisible = True
        Exit For
        End If
        End Select
    End If
Next ws"""

content = content.replace(old_code, new_code)

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_184"', 'Attribute VB_Name = "Goren_Claude_V2_185"')
content = content.replace('VERSION: V2.184', 'VERSION: V2.185')
content = content.replace('APP_VERSION As String = "2.184"', 'APP_VERSION As String = "2.185"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.185 correctly!")
