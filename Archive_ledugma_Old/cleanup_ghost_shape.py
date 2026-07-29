import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Add shpDismissMsg to shapesToDelete so it gets cleaned up
    pattern = re.compile(r'shapesToDelete = Array\(.*?\)w*', re.DOTALL)
    
    # Alternatively, just add it to the explicit delete section
    if 'wsMain.Shapes("shpDismissMsg").Delete' not in content[:6000]:
        # Let's insert it around line 4040 where shapesToDelete loop is
        pattern2 = re.compile(r'(For Each sName In shapesToDelete\n.*?Next sName\n)')
        if pattern2.search(content):
            content = pattern2.sub(r'\1    On Error Resume Next\n    wsMain.Shapes("shpDismissMsg").Delete\n    On Error GoTo 0\n', content)
            with open(bas_file, 'w', encoding='utf-8') as f:
                f.write(content)
            print("Added shpDismissMsg cleanup.")
        else:
            print("Could not find shapesToDelete loop.")
    else:
        print("Cleanup already exists.")

if __name__ == "__main__":
    main()
