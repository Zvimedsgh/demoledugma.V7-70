import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.31.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # Move M23 to M21
    content = content.replace('wsMain.Range("M23")', 'wsMain.Range("M21")')
    
    # Move A23 to A21
    content = content.replace('wsMain.Range("A23")', 'wsMain.Range("A21")')
    
    # Move the old R1:T23 clear range to R1:T21 just to be neat
    content = content.replace('wsMain.Range("R1:T23").ClearContents', 'wsMain.Range("R1:T21").ClearContents')

    # Update comments referencing 23
    content = content.replace('Add Credit and Version to C23', 'Add Credit and Version to A21')
    content = content.replace('Show/Hide Sheets in S23', 'Show/Hide Sheets in M21')

    # Bump version
    content = content.replace('Attribute VB_Name = "Goren_Claude1.31"', 'Attribute VB_Name = "Goren_Claude1.32"')
    content = content.replace("' VERSION: V1.31", "' VERSION: V1.32")
    content = content.replace('Private Const APP_VERSION As String = "1.31"', 'Private Const APP_VERSION As String = "1.32"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
