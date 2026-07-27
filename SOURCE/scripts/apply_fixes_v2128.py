import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.127.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_demo_msg = """Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D15").Left, wsMain.Range("D15").Top, 350, 40)
shpDemoMsg.Name = "shpDemoMsgText"
shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
shpDemoMsg.TextFrame2.TextRange.Font.Size = 17"""

new_demo_msg = """Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D15").Left, wsMain.Range("D15").Top, 500, 40)
shpDemoMsg.Name = "shpDemoMsgText"
shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " - 054-6677396"
shpDemoMsg.TextFrame2.TextRange.Font.Size = 17"""

if target_demo_msg in content:
    content = content.replace(target_demo_msg, new_demo_msg)
    print("Replaced demo text.")
else:
    print("Could not find target_demo_msg")


target_demo_msg2 = """3685 Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D15").Left, wsMain.Range("D15").Top, 350, 40)
3686 shpDemoMsg.Name = "shpDemoMsgText"
3687 shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
3688 shpDemoMsg.TextFrame2.TextRange.Font.Size = 17"""

new_demo_msg2 = """3685 Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D15").Left, wsMain.Range("D15").Top, 500, 40)
3686 shpDemoMsg.Name = "shpDemoMsgText"
3687 shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " - 054-6677396"
3688 shpDemoMsg.TextFrame2.TextRange.Font.Size = 17"""

if target_demo_msg2 in content:
    content = content.replace(target_demo_msg2, new_demo_msg2)
    print("Replaced demo text in setup.")
else:
    print("Could not find target_demo_msg2")


content = content.replace('Attribute VB_Name = "Goren_Claude_V2_127"', 'Attribute VB_Name = "Goren_Claude_V2_128"')
content = content.replace('VERSION: V2.127', 'VERSION: V2.128')
content = content.replace('APP_VERSION As String = "2.127"', 'APP_VERSION As String = "2.128"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.128.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.128 created.")
