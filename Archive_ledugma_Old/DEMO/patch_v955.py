def create_v955():
    with open(r'C:\ledugma\DEMO\V9.54_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    # 1. Delete Credit from A16
    code = code.replace(
        'wsMain.Range("A16").Value = ChrW(1508) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1495)',
        'wsMain.Range("A16").ClearContents \' Removed redundant credit\n        \' wsMain.Range("A16").Value = ChrW(1508)'
    )
    # Remove the formatting lines for A16 to avoid errors or leaving bold text
    code = code.replace('wsMain.Range("A16").Font.Size = 12', '')
    code = code.replace('wsMain.Range("A16").Font.Bold = True', '')
    code = code.replace('wsMain.Range("A16").Font.Color = RGB(0, 100, 0)', '')
    code = code.replace('wsMain.Range("A16").HorizontalAlignment = xlLeft', '')
    code = code.replace('wsMain.Range("A16").VerticalAlignment = xlCenter', '')

    # 2. Widen column G
    # Let's add it near row height settings
    code = code.replace(
        'wsMain.Rows("1:1").RowHeight = 50',
        'wsMain.Rows("1:1").RowHeight = 50\n    wsMain.Columns("G:G").ColumnWidth = 25'
    )

    # 3. Put default text in G11 (הפעל כפתור חפש)
    # We can add this near where F11 "שם לקוח" is set
    f11_code = 'wsMain.Range("F11").Value = ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)'
    g11_default = 'wsMain.Range("G11").Value = ChrW(1492) & ChrW(1508) & ChrW(1506) & ChrW(1500) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " " & ChrW(1495) & ChrW(1508) & ChrW(1513)'
    code = code.replace(f11_code, f11_code + '\n        ' + g11_default)

    # Note: G7 not updating is due to the user not pasting ThisWorkbook code.
    # We will explain this in the response.

    with open(r'C:\ledugma\DEMO\V9.55_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v955()
