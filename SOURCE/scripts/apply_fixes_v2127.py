import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.126.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_search_expl = """' Add explanation text
wsSearch.Cells(5, 2).Value = ChrW(1492) & ChrW(1505) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & ":"
wsSearch.Cells(6, 2).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1512) & ChrW(1511) & " 77"
wsSearch.Cells(7, 2).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " *77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1500) & " 77"
wsSearch.Cells(8, 2).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & "-77"
wsSearch.Cells(9, 2).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " ???: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1506) & ChrW(1501) & " 3 " & ChrW(1514) & ChrW(1493) & ChrW(1493) & ChrW(1497) & ChrW(1501)
wsSearch.Range("B5:B9").Font.Size = 11
wsSearch.Range("B5:B9").Font.Color = RGB(100, 100, 100)
wsSearch.Cells(5, 2).Font.Bold = True"""

new_search_expl = """' Add explanation text
wsSearch.Cells(5, 5).Value = ChrW(1492) & ChrW(1505) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & ":"
wsSearch.Cells(6, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1512) & ChrW(1511) & " 77"
wsSearch.Cells(7, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " *77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1500) & " 77"
wsSearch.Cells(8, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & "-77"
wsSearch.Cells(9, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " ???: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1506) & ChrW(1501) & " 3 " & ChrW(1514) & ChrW(1493) & ChrW(1493) & ChrW(1497) & ChrW(1501)
wsSearch.Range("E5:E9").Font.Size = 12
wsSearch.Range("E5:E9").Font.Color = RGB(0, 112, 192)
wsSearch.Range("E5:E9").HorizontalAlignment = -4152 ' xlRight
wsSearch.Cells(5, 5).Font.Bold = True"""

if target_search_expl in content:
    content = content.replace(target_search_expl, new_search_expl)
    print("Replaced search text.")
else:
    print("Could not find target_search_expl")

target_demo_msg = """Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D3").Left, wsMain.Range("D3").Top, 350, 40)
shpDemoMsg.Name = "shpDemoMsgText"
shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
shpDemoMsg.TextFrame2.TextRange.Font.Size = 16"""

new_demo_msg = """Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D15").Left, wsMain.Range("D15").Top, 350, 40)
shpDemoMsg.Name = "shpDemoMsgText"
shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
shpDemoMsg.TextFrame2.TextRange.Font.Size = 17"""

if target_demo_msg in content:
    content = content.replace(target_demo_msg, new_demo_msg)
    print("Replaced demo text.")
else:
    print("Could not find target_demo_msg")


content = content.replace('Attribute VB_Name = "Goren_Claude_V2_126"', 'Attribute VB_Name = "Goren_Claude_V2_127"')
content = content.replace('VERSION: V2.126', 'VERSION: V2.127')
content = content.replace('APP_VERSION As String = "2.126"', 'APP_VERSION As String = "2.127"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.127.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.127 created.")
