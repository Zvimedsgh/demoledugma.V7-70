with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Orit_Final.bas', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Zoom to 1.2
text = text.replace('newZoom = CLng(ActiveWindow.Zoom * 1.3)', 'newZoom = CLng(ActiveWindow.Zoom * 1.2)')

# 2. RowHeight = 90
text = text.replace('wsMain.Rows("1").RowHeight = 120', 'wsMain.Rows("1").RowHeight = 90')

# 3. Restore 'גורנטק' instead of 'גירסה'
# The current text is ChrW(1490) & ChrW(1497) & ChrW(1512) & ChrW(1505) & ChrW(1492)
# I want ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511)
gorentech = 'ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511)'
girsa = 'ChrW(1490) & ChrW(1497) & ChrW(1512) & ChrW(1505) & ChrW(1492)'

text = text.replace(girsa, gorentech)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Orit_Final.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print('Edits applied')
