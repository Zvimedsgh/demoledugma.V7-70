import re

def create_v953():
    with open(r'C:\ledugma\DEMO\V9.52_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    # 1. Enable events at the start of A00
    code = code.replace(
        'Public Sub SetupMainSheet()\n',
        'Public Sub SetupMainSheet()\n    Application.EnableEvents = True\n'
    )

    # 2. Fix the overlapping buttons under F and G
    # We will replace the entire blocks for btnSearch, btnAll, and btnClearSel.
    # To do this safely, we will find the block starting from 'Set btnSearch =' to 'btnClearSel.OnAction = "ClearSelection"'
    # and replace the positioning logic.
    
    # We will use regex to find and replace the AddShape lines.
    
    # btnSearch
    code = re.sub(
        r'Set btnSearch = wsMain\.Shapes\.AddShape\(msoShapeRoundedRectangle, _[^)]+\)',
        'Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F13").Left + (wsMain.Range("F13").Width - 70) / 2, btnSearchTop, 70, 30)',
        code, flags=re.MULTILINE
    )
    
    # btnAll
    code = re.sub(
        r'Set btnAll = wsMain\.Shapes\.AddShape\(msoShapeRoundedRectangle, _[^)]+\)',
        'Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G13").Left + (wsMain.Range("G13").Width - 70) / 2, btnSearchTop, 70, 30)',
        code, flags=re.MULTILINE
    )
    
    # Fix btnClearTop calculation to avoid overlap
    code = code.replace(
        'btnClearTop = btnSearchTop + 25 + 5',
        'btnClearTop = btnSearchTop + 30 + 10'
    )
    
    # btnClearSel
    code = re.sub(
        r'Set btnClearSel = wsMain\.Shapes\.AddShape\(msoShapeRoundedRectangle, _[^)]+\)',
        'Set btnClearSel = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F13:G13").Left + (wsMain.Range("F13:G13").Width - 150) / 2, btnClearTop, 150, 30)',
        code, flags=re.MULTILINE
    )

    # 3. Move Credit from A17 to A14
    code = code.replace(
        'wsMain.Range("A17").Value = ChrW(1508) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1495)',
        'wsMain.Range("A14").Value = ChrW(1508) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1495)'
    )
    code = code.replace('wsMain.Range("A17").Font.Size = 12', 'wsMain.Range("A14").Font.Size = 12')
    code = code.replace('wsMain.Range("A17").Font.Bold = True', 'wsMain.Range("A14").Font.Bold = True')
    code = code.replace('wsMain.Range("A17").Font.Color = RGB(0, 100, 0)', 'wsMain.Range("A14").Font.Color = RGB(0, 100, 0)')
    code = code.replace('wsMain.Range("A17").HorizontalAlignment = xlLeft', 'wsMain.Range("A14").HorizontalAlignment = xlLeft')
    code = code.replace('wsMain.Range("A17").VerticalAlignment = xlCenter', 'wsMain.Range("A14").VerticalAlignment = xlCenter')

    # Also move processing message from F17 to F15
    code = code.replace('With wsMain.Range("F17")', 'With wsMain.Range("F15")')

    with open(r'C:\ledugma\DEMO\V9.53_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v953()
