import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.130.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update A00_SetupMainSheet demo msg text & width
target_setup = """Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D3").Left, wsMain.Range("D3").Top, 350, 40)
shpDemoMsg.Name = "shpDemoMsgText"
shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
shpDemoMsg.TextFrame2.TextRange.Font.Size = 16"""

# actually wait, I know the exact lines in V2.130 from my previous checks
target_setup_v2130 = """3685 Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D15").Left, wsMain.Range("D15").Top, 350, 40)
3686 shpDemoMsg.Name = "shpDemoMsgText"
3687 shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
3688 shpDemoMsg.TextFrame2.TextRange.Font.Size = 17"""

new_setup_v2130 = """3685 Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D15").Left, wsMain.Range("D15").Top, 500, 40)
3686 shpDemoMsg.Name = "shpDemoMsgText"
3687 shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " - 054-6677396"
3688 shpDemoMsg.TextFrame2.TextRange.Font.Size = 17"""

if target_setup_v2130 in content:
    content = content.replace(target_setup_v2130, new_setup_v2130)
    print("Replaced Setup demo msg.")
else:
    print("Could not find Setup demo msg.")


# 2. Move Explanation block BEFORE Button block
expl_block = """' Add explanation text
wsSearch.Cells(5, 5).Value = ChrW(1492) & ChrW(1505) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & ":"
wsSearch.Cells(6, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1512) & ChrW(1511) & " 77"
wsSearch.Cells(7, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " *77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1500) & " 77"
wsSearch.Cells(8, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & "-77"
wsSearch.Cells(9, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " ???: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1506) & ChrW(1501) & " 3 " & ChrW(1514) & ChrW(1493) & ChrW(1493) & ChrW(1497) & ChrW(1501)
wsSearch.Range("E5:E9").Font.Size = 12
wsSearch.Range("E5:E9").Font.Color = RGB(0, 112, 192)
wsSearch.Cells(5, 5).Font.Bold = True"""

btn_start = "' Add \"Search\" button"

if expl_block in content and btn_start in content:
    content = content.replace(expl_block + "\n", "")
    content = content.replace(expl_block, "")
    
    new_btn_block = expl_block + "\n\n" + btn_start
    content = content.replace(btn_start, new_btn_block)
    print("Moved explanation before buttons.")
else:
    print("Could not find blocks.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_130"', 'Attribute VB_Name = "Goren_Claude_V2_131"')
content = content.replace('VERSION: V2.130', 'VERSION: V2.131')
content = content.replace('APP_VERSION As String = "2.130"', 'APP_VERSION As String = "2.131"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.131.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.131 created.")
