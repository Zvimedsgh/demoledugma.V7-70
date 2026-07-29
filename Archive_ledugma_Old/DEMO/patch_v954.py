def create_v954():
    with open(r'C:\ledugma\DEMO\V9.53_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    # Move Credit back to A16
    code = code.replace(
        'wsMain.Range("A14").Value = ChrW(1508) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1495)',
        'wsMain.Range("A16").Value = ChrW(1508) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1495)'
    )
    code = code.replace('wsMain.Range("A14").Font.Size = 12', 'wsMain.Range("A16").Font.Size = 12')
    code = code.replace('wsMain.Range("A14").Font.Bold = True', 'wsMain.Range("A16").Font.Bold = True')
    code = code.replace('wsMain.Range("A14").Font.Color = RGB(0, 100, 0)', 'wsMain.Range("A16").Font.Color = RGB(0, 100, 0)')
    code = code.replace('wsMain.Range("A14").HorizontalAlignment = xlLeft', 'wsMain.Range("A16").HorizontalAlignment = xlLeft')
    code = code.replace('wsMain.Range("A14").VerticalAlignment = xlCenter', 'wsMain.Range("A16").VerticalAlignment = xlCenter')

    # Move processing message back to F17
    code = code.replace('With wsMain.Range("F15")', 'With wsMain.Range("F17")')

    # Reduce Row 2 height and Row 16 height.
    # Currently it is: wsMain.Rows("2:3").RowHeight = 45
    code = code.replace(
        'wsMain.Rows("2:3").RowHeight = 45',
        'wsMain.Rows("3:3").RowHeight = 45\n        wsMain.Rows("2:2").RowHeight = 22.5\n        wsMain.Rows("16:16").RowHeight = 15'
    )

    with open(r'C:\ledugma\DEMO\V9.54_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v954()
