import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.23.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Fix btnSearch
    old_search = 'Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G11").Left + (wsMain.Range("G11").Width - 60) / 2, wsMain.Range("G11").Top + 5, 60, 25)'
    new_search = 'Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F11").Left + (wsMain.Range("F11").Width - 60) / 2, wsMain.Range("F11").Top + 5, 60, 25)'
    content = content.replace(old_search, new_search)

    # 2. Bump version to 1.24
    content = content.replace('Attribute VB_Name = "Goren_Claude1.23"', 'Attribute VB_Name = "Goren_Claude1.24"')
    content = content.replace("' VERSION: V1.23", "' VERSION: V1.24")
    content = content.replace('Private Const APP_VERSION As String = "1.23"', 'Private Const APP_VERSION As String = "1.24"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
