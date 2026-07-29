import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Revert all the bad injections
    bad_code = 'Err.Clear\n    On Error Resume Next\n    wsMain.Shapes("shpDismissMsg").Delete\n    On Error GoTo 0\n'
    content = content.replace(bad_code, 'Err.Clear\n')
    
    # Also some might have different line endings
    bad_code2 = 'Err.Clear\r\n    On Error Resume Next\r\n    wsMain.Shapes("shpDismissMsg").Delete\r\n    On Error GoTo 0\r\n'
    content = content.replace(bad_code2, 'Err.Clear\r\n')

    # Now we inject it PROPERLY into A00_SetupMainSheet only!
    # A00_SetupMainSheet has a shapesToDelete loop around line 4030
    # Let's just find the start of A00_SetupMainSheet
    pattern = re.compile(r'(Public Sub A00_SetupMainSheet\(\)\n.*?Dim shapesToDelete As Variant\n)', re.DOTALL)
    if pattern.search(content):
        # We will inject the deletion right after "Dim shapesToDelete As Variant"
        # Wait, there's already a loop for shapesToDelete that deletes shapes. 
        # I'll just add "shpDismissMsg" to the array!
        # `shapesToDelete = Array("btnBuildReview", "btnApplyCorrections", ... , "shpDemoMsgText")`
        pattern_array = re.compile(r'("btnNavExit", "shpBtn", "shpDemoMsgText"\))')
        if pattern_array.search(content):
            content = pattern_array.sub(r'"btnNavExit", "shpBtn", "shpDemoMsgText", "shpDismissMsg")', content)
            print("Successfully added shpDismissMsg to shapesToDelete.")
        else:
            print("Could not find shapesToDelete array.")
    else:
        print("Could not find A00_SetupMainSheet.")

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
