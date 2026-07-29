import re

def create_v952():
    with open(r'C:\ledugma\DEMO\V9.51_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    # 1. Fix row 4 height
    code = code.replace('wsMain.Rows("2:4").RowHeight = 45', 'wsMain.Rows("2:3").RowHeight = 45\n        wsMain.Rows("4:11").RowHeight = 25')
    
    # 2. Fix 6 main buttons height. They were squeezed because the parameter box is shorter now.
    # We will remove the dynamic height calculation that tied them to the Clear button bottom.
    code = re.sub(
        r'btnHeight = \(totalAvailHeight - \(5 \* btnGap\)\) / 6',
        'btnHeight = 35 \' Fixed height instead of squeezing',
        code
    )
    
    # 3. Clear old G12:G13 leftovers
    code = code.replace('wsMain.Range("A3:D6").ClearContents', 'wsMain.Range("A3:D6").ClearContents\n        wsMain.Range("F12:G15").Clear \' Remove old leftovers like F13/G13')
    
    # 4. Make search and clear buttons equal size and centered
    # They are located relative to F14/G14. But wait, we just cleared F12:G15! That's fine, we use their Left/Width properties.
    # The width of F14 is column F width. F14:G14 is F+G width.
    # Let's just use F13:G13 for centering to be safe.
    # width = 120, height = 30 for BOTH.
    
    # Replace search button sizing
    code = re.sub(
        r'wsMain\.Range\("F14"\)\.Left \+ \(wsMain\.Range\("F14"\)\.Width - \d+\) / 2, _\s+btnSearchTop, \d+, \d+\)',
        'wsMain.Range("F13:G13").Left + (wsMain.Range("F13:G13").Width - 120) / 2, _\n            btnSearchTop, 120, 30)',
        code
    )
    
    # Replace clear button sizing
    code = re.sub(
        r'wsMain\.Range\("F15:G15"\)\.Left \+ \(wsMain\.Range\("F15:G15"\)\.Width - \d+\) / 2, _\s+btnClearTop, \d+, \d+\)',
        'wsMain.Range("F13:G13").Left + (wsMain.Range("F13:G13").Width - 120) / 2, _\n            btnClearTop, 120, 30)',
        code
    )
    
    # 5. Fix G10 conditional format to ensure it uses FormatConditions(1) just like G7
    # Look for G10 format condition color
    code = code.replace(
        'wsMain.Range("G10").FormatConditions(wsMain.Range("G10").FormatConditions.count).Interior.Color = RGB(255, 215, 0)',
        'wsMain.Range("G10").FormatConditions(1).Interior.Color = RGB(255, 215, 0)'
    )

    with open(r'C:\ledugma\DEMO\V9.52_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v952()
