import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.24.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. btnReset positioning
    old_reset = 'Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G12").Left + (wsMain.Range("G12").Width + wsMain.Range("F12").Width - 120) / 2, wsMain.Range("G12").Top + 5, 120, 25)'
    new_reset = 'Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G11").Left + (wsMain.Range("G11").Width + wsMain.Range("F11").Width - 120) / 2, wsMain.Range("F11").Top + 35, 120, 25)'
    content = content.replace(old_reset, new_reset)

    # 2. Credit text
    content = content.replace('wsMain.Range("A22").Value = ', 'wsMain.Range("A23").Value = ')
    content = content.replace('wsMain.Range("A22").Font.Color', 'wsMain.Range("A23").Font.Color')
    content = content.replace('wsMain.Range("A22").Font.Size', 'wsMain.Range("A23").Font.Size')
    content = content.replace('wsMain.Range("A22").Font.Bold', 'wsMain.Range("A23").Font.Bold')
    content = content.replace('wsMain.Range("A22").HorizontalAlignment = xlCenter', 'wsMain.Range("A23").HorizontalAlignment = xlLeft')
    
    # Remove A22:D24 clear contents so A23 isn't deleted
    content = content.replace('wsMain.Range("A22:D24").ClearContents', "' Removed ClearContents to protect credit text in A23")

    # 3. ToggleHiddenSheets
    content = content.replace('wsMain.Range("P22").Left, wsMain.Range("P22").Top, 160, 30', 'wsMain.Range("M23").Left, wsMain.Range("M23").Top, 160, 30')

    # 4. Matach text size
    content = content.replace('wsMain.Range("J5").Font.Size = 9', 'wsMain.Range("J5").Font.Size = 7\n    wsMain.Range("J5").ShrinkToFit = True')

    # Bump version
    content = content.replace('Attribute VB_Name = "Goren_Claude1.24"', 'Attribute VB_Name = "Goren_Claude1.25"')
    content = content.replace("' VERSION: V1.24", "' VERSION: V1.25")
    content = content.replace('Private Const APP_VERSION As String = "1.24"', 'Private Const APP_VERSION As String = "1.25"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
