import os

def fix_labels_and_buttons():
    with open(r'C:\ledugma\DEMO\V9.50_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    # 1. Increase button width
    code = code.replace(
        'wsMain.Range("F14").Left + (wsMain.Range("F14").Width - 50) / 2, _\n            btnSearchTop, 50, 25)',
        'wsMain.Range("F14").Left + (wsMain.Range("F14").Width - 100) / 2, _\n            btnSearchTop, 100, 30)'
    )
    
    code = code.replace(
        'wsMain.Range("F15:G15").Left + (wsMain.Range("F15:G15").Width - 105) / 2, _\n            btnClearTop, 105, 25)',
        'wsMain.Range("F15:G15").Left + (wsMain.Range("F15:G15").Width - 130) / 2, _\n            btnClearTop, 130, 30)'
    )
    
    # 2. Add labels explicitly
    label_code = """
        ' explicitly set F labels
        wsMain.Range("F4").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1497) & ChrW(1497) & ChrW(1495) & ChrW(1493) & ChrW(1505)
        wsMain.Range("F5").Value = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1502) & ChrW(1499) & ChrW(1508) & ChrW(1497) & ChrW(1500)
        wsMain.Range("F6").Value = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
        wsMain.Range("F7").Value = ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
        
        wsMain.Range("F9").Value = ChrW(1495) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1498) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497)
        wsMain.Range("F10").Value = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1495) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1498)
    """
    
    # inject after clear A3:D6
    target = '130     wsMain.Range("A3:D6").ClearContents'
    code = code.replace(target, target + '\n' + label_code)
    
    # Also remove AGENCY_NAME from presentation title
    code = code.replace('shp.TextFrame.TextRange.Text = AGENCY_NAME()', 'shp.TextFrame.TextRange.Text = ""')
    
    # 3. Add safety in G7 event?
    # Actually just add it to V9.51
    with open(r'C:\ledugma\DEMO\V9.51_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    fix_labels_and_buttons()
