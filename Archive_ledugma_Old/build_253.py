import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.251.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.251', 'VERSION: V2.253')
content = content.replace('Error in V2.251!', 'Error in V2.253!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_251"', 'Attribute VB_Name = "Goren_Claude_V2_253"')
content = content.replace('APP_VERSION As String = "2.251"', 'APP_VERSION As String = "2.253"')
content = content.replace('APP_VERSION = "2.251"', 'APP_VERSION = "2.253"')

insertion_point = content.find("' ---- Set matach sheet tab color (brown) and hide it ----")

credit_code = """
    ' ---- Add WhatsApp Credit in Row 19 ----
4791 On Error Resume Next
4792 wsMain.Range("C19:K20").UnMerge
4793 With wsMain.Range("C19:I19")
4794     .MergeCells = True
4795     .Value = ChrW(1500) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1493) & ChrW(1505) & ChrW(1497) & ChrW(1493) & ChrW(1506) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & ChrW(1493) & ChrW(1493) & ChrW(1496) & ChrW(1505) & ChrW(1488) & ChrW(1508) & " " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & " 054-6677396"
4796     .Font.Size = 16
4797     .Font.Bold = True
4798     .Font.Color = RGB(0, 176, 240) ' Light Blue (tchelet)
4799     .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
    End With
    Err.Clear
    On Error GoTo ERR_HANDLER

"""

new_content = content[:insertion_point] + credit_code + content[insertion_point:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.253.bas', 'w', encoding='utf-8') as f:
    f.write(new_content)
print('Generated 253')
