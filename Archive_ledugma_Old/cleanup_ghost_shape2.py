import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    if 'wsMain.Shapes("shpDismissMsg").Delete' not in content[:6000]:
        content = content.replace('Err.Clear', 'Err.Clear\n    On Error Resume Next\n    wsMain.Shapes("shpDismissMsg").Delete\n    On Error GoTo 0\n')
        with open(bas_file, 'w', encoding='utf-8') as f:
            f.write(content)
        print("Added shpDismissMsg cleanup.")
    else:
        print("Cleanup already exists.")

if __name__ == "__main__":
    main()
