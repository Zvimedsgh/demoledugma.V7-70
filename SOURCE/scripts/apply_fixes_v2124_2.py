import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.124.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_reset_home = """wsMain.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)
wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)"""
new_reset_home = """wsMain.Range("G10").Value = ""
wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)
On Error Resume Next
wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513) ' "חפש"
On Error GoTo 0"""

if target_reset_home in content:
    content = content.replace(target_reset_home, new_reset_home)
    print("Replaced target_reset_home")
else:
    print("Could not find target_reset_home")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

