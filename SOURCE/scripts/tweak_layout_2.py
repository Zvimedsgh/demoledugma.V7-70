import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.21.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Row 1 height 60
    content = content.replace('wsMain.Rows("1:24").RowHeight = 22', 'wsMain.Rows("1:24").RowHeight = 22\n    wsMain.Rows("1").RowHeight = 60')
    content = content.replace('wsMain.Rows("11:13").RowHeight = 35', '') # Remove the 35 height since it breaks button 6 gap

    # 2. Columns H:I width 13 each
    content = content.replace('wsMain.Columns("H").ColumnWidth = 8', 'wsMain.Columns("H").ColumnWidth = 13')
    content = content.replace('wsMain.Columns("I").ColumnWidth = 8', 'wsMain.Columns("I").ColumnWidth = 13')

    # 3. Center Search button in F11 (Right column)
    content = content.replace('wsMain.Range("F12").Left + 10', 'wsMain.Range("F11").Left + (wsMain.Range("F11").Width - 60) / 2')
    content = content.replace('wsMain.Range("F12").Top + 10', 'wsMain.Range("F11").Top + 5')

    # 4. Center All button in G11 (Left column)
    content = content.replace('wsMain.Range("G12").Left + 10', 'wsMain.Range("G11").Left + (wsMain.Range("G11").Width - 60) / 2')
    content = content.replace('wsMain.Range("G12").Top + 10', 'wsMain.Range("G11").Top + 5')

    # 5. Center Reset button in F12:G12
    old_reset_pos = 'Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G14").Left + 20, wsMain.Range("G14").Top + 5, (wsMain.Range("G14").Width + wsMain.Range("F14").Width) - 40, 30)'
    new_reset_pos = 'Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G12").Left + (wsMain.Range("G12").Width + wsMain.Range("F12").Width - 120) / 2, wsMain.Range("G12").Top + 5, 120, 25)'
    content = content.replace(old_reset_pos, new_reset_pos)

    # 6. Move credit to A22, start in last third
    content = content.replace('wsMain.Range("C23").Value', 'wsMain.Range("A22").Value')
    content = content.replace('wsMain.Range("C23").Font.Color', 'wsMain.Range("A22").Font.Color')
    content = content.replace('wsMain.Range("C23").Font.Size', 'wsMain.Range("A22").Font.Size')
    content = content.replace('wsMain.Range("C23").Font.Bold', 'wsMain.Range("A22").Font.Bold')
    content = content.replace('wsMain.Range("C23").HorizontalAlignment = xlCenter', 'wsMain.Range("A22").HorizontalAlignment = xlLeft')

    # 7. Move Show/Hide sheets button to P22
    content = content.replace('wsMain.Range("S23").Left', 'wsMain.Range("P22").Left')
    content = content.replace('wsMain.Range("S23").Top', 'wsMain.Range("P22").Top')

    # 8. Fix button 5 and 6 gap
    # Because we removed the height=35 for row 11, we don't necessarily need to hardcode the Top if we just use C12.
    # Wait, the buttons are on C2, C4, C6, C8, C10. If the next is C12, the gap is C11.
    # Since C11 is now 22 height, the gap is identical to C3, C5, C7, C9.
    # So we DO NOT need to hardcode it anymore! C12 is perfectly correct.
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
