import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038_20260702_1440.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_block = """wsMain.Range("C1:K1").Merge
wsMain.Range("C1").Value = GetActiveAgencyName()
wsMain.Range("C1").Font.Size = 28
wsMain.Range("C1").Font.Bold = True
wsMain.Range("C1").Font.Color = RGB(200, 0, 0)
wsMain.Range("C1").HorizontalAlignment = xlCenter
wsMain.Range("C1").VerticalAlignment = xlCenter"""

new_block = """wsMain.Range("D1:K1").Merge
wsMain.Range("D1").Value = GetActiveAgencyName()
wsMain.Range("D1").Font.Size = 28
wsMain.Range("D1").Font.Bold = True
wsMain.Range("D1").Font.Color = RGB(200, 0, 0)
wsMain.Range("D1").HorizontalAlignment = xlCenter
wsMain.Range("D1").VerticalAlignment = xlCenter"""

content = content.replace(old_block, new_block)
content = content.replace('Private Const APP_VERSION As String = "2.038"', 'Private Const APP_VERSION As String = "2.039"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.039.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("Title shifted left by changing merge area to D1:K1")
