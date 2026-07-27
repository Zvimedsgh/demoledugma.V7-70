import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.131.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Search Explanation
old_search_expl = """wsSearch.Cells(8, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & "-77"
wsSearch.Cells(9, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " ???: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1506) & ChrW(1501) & " 3 " & ChrW(1514) & ChrW(1493) & ChrW(1493) & ChrW(1497) & ChrW(1501)
wsSearch.Range("E5:E9").Font.Size = 12
wsSearch.Range("E5:E9").Font.Color = RGB(0, 112, 192)"""

new_search_expl = """wsSearch.Cells(8, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & "-77"
wsSearch.Range("E5:E8").Font.Size = 12
wsSearch.Range("E5:E8").Font.Color = RGB(0, 112, 192)"""

content = content.replace(old_search_expl, new_search_expl)


# 2. Demo message position (D15 -> F19)
content = content.replace('wsMain.Range("D15").Left, wsMain.Range("D15").Top', 'wsMain.Range("F19").Left, wsMain.Range("F19").Top')


# 3. Credit text (F19 -> A21)
old_credit = """3970 wsMain.Range("F19").Value = ChrW(1504) & ChrW(1489) & ChrW(1504) & ChrW(1492) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1497) & ChrW(1491) & ChrW(1497) & " " & ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " 054-6677396 - v" & APP_VERSION
3980 wsMain.Range("F19").Font.Size = 12
3990 wsMain.Range("F19").Font.Color = RGB(0, 0, 0)
4000 wsMain.Range("F19").Font.Bold = True
4010 wsMain.Range("F19").HorizontalAlignment = -4108"""

new_credit = """3970 wsMain.Range("A22").Value = ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " v" & APP_VERSION
3980 wsMain.Range("A22").Font.Size = 10
3990 wsMain.Range("A22").Font.Color = RGB(150, 150, 150)
4000 wsMain.Range("A22").Font.Bold = False
4010 wsMain.Range("A22").HorizontalAlignment = -4108"""

# I need to clear A22 explicitly or maybe not, the old F19 will be overwritten by the Demo msg (or we should clear F19? No, the Demo msg is a Shape, it doesn't clear the cell text!).
# Wait! I need to clear F19 so it doesn't show the old text!
new_credit_full = """3970 wsMain.Range("F19").ClearContents
3975 wsMain.Range("A22").Value = ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " v" & APP_VERSION
3980 wsMain.Range("A22").Font.Size = 10
3990 wsMain.Range("A22").Font.Color = RGB(150, 150, 150)
4000 wsMain.Range("A22").Font.Bold = False
4010 wsMain.Range("A22").HorizontalAlignment = -4108"""

content = content.replace(old_credit, new_credit_full)

# Rename to V2.132
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_131"', 'Attribute VB_Name = "Goren_Claude_V2_132"')
content = content.replace('VERSION: V2.131', 'VERSION: V2.132')
content = content.replace('APP_VERSION As String = "2.131"', 'APP_VERSION As String = "2.132"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.132.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.132 created.")
