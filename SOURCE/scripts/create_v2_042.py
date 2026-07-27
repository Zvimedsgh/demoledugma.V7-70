import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.041_20260702_1520.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Fix title to D1:K1 and add row height
old_title = """wsMain.Range("G1:P1").Merge
wsMain.Range("G1").Value = GetActiveAgencyName()
wsMain.Range("G1").Font.Size = 28
wsMain.Range("G1").Font.Bold = True
wsMain.Range("G1").Font.Color = RGB(200, 0, 0)
wsMain.Range("G1").HorizontalAlignment = xlCenter
wsMain.Range("G1").VerticalAlignment = xlCenter"""

new_title = """wsMain.Range("D1:K1").Merge
wsMain.Range("D1").Value = GetActiveAgencyName()
wsMain.Range("D1").Font.Size = 28
wsMain.Range("D1").Font.Bold = True
wsMain.Range("D1").Font.Color = RGB(200, 0, 0)
wsMain.Range("D1").HorizontalAlignment = xlCenter
wsMain.Range("D1").VerticalAlignment = xlCenter
wsMain.Rows(1).RowHeight = 120"""

content = content.replace(old_title, new_title)

# 2. Add ApplyDemoLockOnOpen before Exit Sub
old_exit = """4085  Application.EnableEvents = True
4090  MsgBoxU ThisWorkbook.Worksheets(H_SET_MESSAGES()).Cells(12, 1).Value, vbInformation
4100  Exit Sub"""

new_exit = """4085  Application.EnableEvents = True
4088  ApplyDemoLockOnOpen
4090  MsgBoxU ThisWorkbook.Worksheets(H_SET_MESSAGES()).Cells(12, 1).Value, vbInformation
4100  Exit Sub"""

content = content.replace(old_exit, new_exit)

# 3. Update version to 2.042
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_041"', 'Attribute VB_Name = "Goren_Claude_V2_042"')
content = content.replace('Private Const APP_VERSION As String = "2.041"', 'Private Const APP_VERSION As String = "2.042"')
content = content.replace('VERSION: V2.041', 'VERSION: V2.042')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.042.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.042 with row height, title shift, and ApplyDemoLockOnOpen")
