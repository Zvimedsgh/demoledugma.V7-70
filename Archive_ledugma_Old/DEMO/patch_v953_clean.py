import re

def fix_v953():
    with open(r'C:\ledugma\DEMO\V9.52_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    # Enable events at start of Setup
    code = code.replace(
        'Public Sub SetupMainSheet()\n',
        'Public Sub SetupMainSheet()\n    Application.EnableEvents = True\n'
    )

    # 1. Fix btnSearch
    # In V9.52:
    # Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, _
    #        wsMain.Range("F13:G13").Left + (wsMain.Range("F13:G13").Width - 120) / 2, _
    #        btnSearchTop, 120, 30)
    search_orig = 'Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, _\n            wsMain.Range("F13:G13").Left + (wsMain.Range("F13:G13").Width - 120) / 2, _\n            btnSearchTop, 120, 30)'
    search_new = 'Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F13").Left + (wsMain.Range("F13").Width - 70) / 2, btnSearchTop, 70, 30)'
    code = code.replace(search_orig, search_new)

    # 2. Fix btnAll
    # In V9.52:
    # Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, _
    #        wsMain.Range("G14").Left + (wsMain.Range("G14").Width - 50) / 2, _
    #        btnSearchTop, 50, 25)
    all_orig = 'Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, _\n            wsMain.Range("G14").Left + (wsMain.Range("G14").Width - 50) / 2, _\n            btnSearchTop, 50, 25)'
    all_new = 'Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G13").Left + (wsMain.Range("G13").Width - 70) / 2, btnSearchTop, 70, 30)'
    code = code.replace(all_orig, all_new)

    # 3. Fix btnClearSel
    # In V9.52:
    # Set btnClearSel = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, _
    #        wsMain.Range("F13:G13").Left + (wsMain.Range("F13:G13").Width - 120) / 2, _
    #        btnClearTop, 120, 30)
    clear_orig = 'Set btnClearSel = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, _\n            wsMain.Range("F13:G13").Left + (wsMain.Range("F13:G13").Width - 120) / 2, _\n            btnClearTop, 120, 30)'
    clear_new = 'Set btnClearSel = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F13:G13").Left + (wsMain.Range("F13:G13").Width - 150) / 2, btnClearTop, 150, 30)'
    code = code.replace(clear_orig, clear_new)

    # Fix btnClearTop
    code = code.replace(
        'btnClearTop = btnSearchTop + 25 + 5',
        'btnClearTop = btnSearchTop + 30 + 10'
    )

    # Move Credit
    code = code.replace(
        'wsMain.Range("A17").Value = ChrW(1508) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1495)',
        'wsMain.Range("A14").Value = ChrW(1508) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1495)'
    )
    code = code.replace('wsMain.Range("A17").Font.Size = 12', 'wsMain.Range("A14").Font.Size = 12')
    code = code.replace('wsMain.Range("A17").Font.Bold = True', 'wsMain.Range("A14").Font.Bold = True')
    code = code.replace('wsMain.Range("A17").Font.Color = RGB(0, 100, 0)', 'wsMain.Range("A14").Font.Color = RGB(0, 100, 0)')
    code = code.replace('wsMain.Range("A17").HorizontalAlignment = xlLeft', 'wsMain.Range("A14").HorizontalAlignment = xlLeft')
    code = code.replace('wsMain.Range("A17").VerticalAlignment = xlCenter', 'wsMain.Range("A14").VerticalAlignment = xlCenter')

    code = code.replace('With wsMain.Range("F17")', 'With wsMain.Range("F15")')

    with open(r'C:\ledugma\DEMO\V9.53_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    fix_v953()
