with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.45_CopyPaste.txt', 'r', encoding='windows-1255') as f:
    text = f.read()

# Fix the bug where font size was reset
target = '            wsLoop.Cells.Font.Size = 14\n        Next wsLoop'
replacement = '            wsLoop.Cells.Font.Size = 14\n        Next wsLoop\n\n        \' Restore A1 huge title font size after loop\n        wsMain.Range("A1").Font.Size = 36'
text = text.replace(target, replacement)

text = text.replace('V9.45', 'V9.47')
text = text.replace('9.45', '9.47')

with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.47_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)
