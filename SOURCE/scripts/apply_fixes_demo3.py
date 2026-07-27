import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.122.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

idx = content.find('1650 If isDemoMode Then')
if idx != -1:
    idx2 = content.find('1750 Else', idx)
    if idx2 != -1:
        new_str = """1650 If isDemoMode Then
1660 wsMain.Range("G3").Value = 2024
1670 wsMain.Range("G4").Value = 2025
' Data Validation block removed to avoid Excel bug
1740 wsMain.Range("G3:G4").Interior.Color = RGB(230, 230, 230) ' Gray out visually
    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    On Error GoTo ERR_HANDLER
    Dim shpDemoMsg As Object
    Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D3").Left, wsMain.Range("D3").Top, 350, 40)
    shpDemoMsg.Name = "shpDemoMsgText"
    shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
    shpDemoMsg.TextFrame2.TextRange.Font.Size = 16
    shpDemoMsg.TextFrame2.TextRange.Font.Bold = -1
    shpDemoMsg.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 255)
    shpDemoMsg.Line.Visible = 0
    shpDemoMsg.Fill.Visible = 0
"""
        new_else = """1750 Else
    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    On Error GoTo ERR_HANDLER
"""
        idx3 = content.find('\n', idx2)
        content = content[:idx] + new_str + new_else + content[idx3+1:]
        print("Replaced demo message")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_122"', 'Attribute VB_Name = "Goren_Claude_V2_123"')
content = content.replace('VERSION: V2.122', 'VERSION: V2.123')
content = content.replace('APP_VERSION As String = "2.122"', 'APP_VERSION As String = "2.123"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.123.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.123 created via manual slice.")
