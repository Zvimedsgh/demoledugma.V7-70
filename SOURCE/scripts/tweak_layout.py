import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.1.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Fix the background color issue (A1:U24 green overwriting cream)
    # Remove it from where it is now:
    content = content.replace('wsMain.Range("A1:U24").Interior.Color = RGB(220, 240, 220)', '')
    # And place it right after clearing colors at the top:
    target_clear = 'wsMain.Range("A2:K20").Interior.ColorIndex = xlNone'
    content = content.replace(target_clear, target_clear + '\n    wsMain.Range("A1:U24").Interior.Color = RGB(220, 240, 220)')

    # 2. Delete Sylvan credit
    credit_str = '''    wsMain.Range("C14").Value = ChrW(1508) & ChrW(1493) & ChrW(1514) & ChrW(1495) & " " & ChrW(1506) & ChrW(34) & ChrW(1497) & " " & ChrW(1505) & ChrW(1497) & ChrW(1500) & ChrW(1489) & ChrW(1503) & " " & ChrW(1496) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1500) & ChrW(1493) & ChrW(1490) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " 054-6677396"
    wsMain.Range("C14").Font.Color = RGB(0, 120, 0)
    wsMain.Range("C14").Font.Bold = True
    wsMain.Range("C14").Font.Size = 10
    wsMain.Range("C14").HorizontalAlignment = xlCenter'''
    content = content.replace(credit_str, '')

    # 3. Change "אפס נתונים" to "איפוס נתונים"
    old_reset_text = 'ChrW(1488) & ChrW(1508) & ChrW(1505) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)'
    new_reset_text = 'ChrW(1488) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)'
    content = content.replace(old_reset_text, new_reset_text)

    # 4. Enlarge column C to width 50
    content = content.replace('wsMain.Columns("C").ColumnWidth = 18', 'wsMain.Columns("C").ColumnWidth = 50')

    # 5. Fix placement of Search, All, and Reset buttons
    # Search button
    content = content.replace('wsMain.Range("F11").Left + 5', 'wsMain.Range("F12").Left + 10')
    content = content.replace('wsMain.Range("F11").Top + 2', 'wsMain.Range("F12").Top + 10')
    
    # All button
    content = content.replace('wsMain.Range("G11").Left + 5', 'wsMain.Range("G12").Left + 10')
    content = content.replace('wsMain.Range("G11").Top + 2', 'wsMain.Range("G12").Top + 10')
    
    # Reset button (center it between F and G, below them)
    # The left edge of G (since G is on the left) is Range("G").Left
    # To center it in the combined width of F+G:
    # Width = 100, combined width is wsMain.Range("G13").Width + wsMain.Range("F13").Width
    old_reset_pos = 'Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G12").Left, wsMain.Range("G12").Top + 5, wsMain.Range("G12").Width + wsMain.Range("F12").Width, 25)'
    new_reset_pos = 'Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G14").Left + 20, wsMain.Range("G14").Top + 5, (wsMain.Range("G14").Width + wsMain.Range("F14").Width) - 40, 30)'
    content = content.replace(old_reset_pos, new_reset_pos)

    # Also remove some leftover blank lines that might have appeared
    content = re.sub(r'\n\s*\n\s*\n', '\n\n', content)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
        
    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
