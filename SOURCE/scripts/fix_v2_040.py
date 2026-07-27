import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.039_20260702_1445.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_block = """wsMain.Range("D1:K1").Merge
wsMain.Range("D1").Value = GetActiveAgencyName()
wsMain.Range("D1").Font.Size = 28
wsMain.Range("D1").Font.Bold = True
wsMain.Range("D1").Font.Color = RGB(200, 0, 0)
wsMain.Range("D1").HorizontalAlignment = xlCenter
wsMain.Range("D1").VerticalAlignment = xlCenter"""

new_block = """wsMain.Range("F1:M1").Merge
wsMain.Range("F1").Value = GetActiveAgencyName()
wsMain.Range("F1").Font.Size = 28
wsMain.Range("F1").Font.Bold = True
wsMain.Range("F1").Font.Color = RGB(200, 0, 0)
wsMain.Range("F1").HorizontalAlignment = xlCenter
wsMain.Range("F1").VerticalAlignment = xlCenter"""

content = content.replace(old_block, new_block)

old_view_reports = """Public Sub ViewReportsFolder()
Dim reportsPath As String"""

new_view_reports = """Public Sub ViewReportsFolder()
If CheckDemoLock() Then Exit Sub
Dim reportsPath As String"""

content = content.replace(old_view_reports, new_view_reports)

content = content.replace('Private Const APP_VERSION As String = "2.039"', 'Private Const APP_VERSION As String = "2.040"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.040.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.040 with fixes")
