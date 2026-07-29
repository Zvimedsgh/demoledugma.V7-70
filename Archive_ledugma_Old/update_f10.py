import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Replace Application.Goto instances of F3 with F10
    content = content.replace('Application.Goto wsMain.Range("F3")', 'Application.Goto wsMain.Range("F10")')
    content = content.replace('Application.Goto wsMainUI.Range("F3")', 'Application.Goto wsMainUI.Range("F10")')
    content = content.replace('Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F3")', 'Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F10")')

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Updated F3 to F10 in Goto statements.")

if __name__ == "__main__":
    main()
