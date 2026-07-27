import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.27.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # Decrease font size from 8 to 6
    old_font = 'wsMain.Range("J5").Font.Size = 8'
    new_font = 'wsMain.Range("J5").Font.Size = 6'
    content = content.replace(old_font, new_font)

    # Bump version
    content = content.replace('Attribute VB_Name = "Goren_Claude1.27"', 'Attribute VB_Name = "Goren_Claude1.28"')
    content = content.replace("' VERSION: V1.27", "' VERSION: V1.28")
    content = content.replace('Private Const APP_VERSION As String = "1.27"', 'Private Const APP_VERSION As String = "1.28"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
