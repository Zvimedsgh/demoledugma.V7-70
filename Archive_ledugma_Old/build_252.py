import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.251.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.251', 'VERSION: V2.252')
content = content.replace('Error in V2.251!', 'Error in V2.252!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_251"', 'Attribute VB_Name = "Goren_Claude_V2_252"')
content = content.replace('APP_VERSION As String = "2.251"', 'APP_VERSION As String = "2.252"')
content = content.replace('APP_VERSION = "2.251"', 'APP_VERSION = "2.252"')

# Add credit text in row 19
# We will insert it before: ' ---- Set matach sheet tab color (brown) and hide it ----
insertion_point = content.find("' ---- Set matach sheet tab color (brown) and hide it ----")

credit_code = """
    ' ---- Add WhatsApp Credit in Row 19 ----
    On Error Resume Next
    With wsMain.Range("C19:I19")
        .MergeCells = True
        .Value = ChrW(1500) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1493) & ChrW(1505) & ChrW(1497) & ChrW(1493) & ChrW(1506) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & ChrW(1493) & ChrW(1493) & ChrW(1496) & ChrW(1505) & ChrW(1488) & ChrW(1508) & " " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & " 054-6677396"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(0, 176, 240) ' Light Blue (תכלת)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    On Error GoTo ERR_HANDLER

"""
new_content = content[:insertion_point] + credit_code + content[insertion_point:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.252.bas', 'w', encoding='utf-8') as f:
    f.write(new_content)
print('Generated 252')
