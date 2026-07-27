import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.132.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Typo "חיפוס" -> "חיפוש"
# The string is: ChrW(1492) & ChrW(1505) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & ":"
old_typo = 'ChrW(1492) & ChrW(1505) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & ":"'
new_typo = 'ChrW(1492) & ChrW(1505) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & ":"'
content = content.replace(old_typo, new_typo)

# 2. Add RTL alignment and AutoFit
old_expl_styling = """wsSearch.Range("E5:E8").Font.Size = 12
wsSearch.Range("E5:E8").Font.Color = RGB(0, 112, 192)"""

new_expl_styling = """wsSearch.Range("E5:E8").Font.Size = 12
wsSearch.Range("E5:E8").Font.Color = RGB(0, 112, 192)
wsSearch.Range("E5:E8").HorizontalAlignment = -4152 ' xlRight
wsSearch.Columns(5).ColumnWidth = 60 ' Expand column"""

content = content.replace(old_expl_styling, new_expl_styling)

# 3. Two spaces after Whatsapp
old_wa = '" Whatsapp " & ChrW(1500) & " - 054-6677396"'
new_wa = '" Whatsapp  " & ChrW(1500) & " - 054-6677396"'
content = content.replace(old_wa, new_wa)


# Rename to V2.133
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_132"', 'Attribute VB_Name = "Goren_Claude_V2_133"')
content = content.replace('VERSION: V2.132', 'VERSION: V2.133')
content = content.replace('APP_VERSION As String = "2.132"', 'APP_VERSION As String = "2.133"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.133.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.133 created.")
