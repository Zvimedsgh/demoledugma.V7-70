import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Update ApplyDemoLockOnOpen signature
    content = content.replace('Public Sub ApplyDemoLockOnOpen()', 'Public Sub ApplyDemoLockOnOpen(Optional ByVal isSetup As Boolean = False)')

    # 2. Update the call in A00_SetupMainSheet
    content = content.replace('4940 ApplyDemoLockOnOpen\n', '4940 ApplyDemoLockOnOpen True\n')

    # 3. Wrap the stripping logic in ApplyDemoLockOnOpen with `If Not isSetup Then`
    # The stripping logic starts with `' Desktop: Remove Hyperlink` and ends with `On Error GoTo 0`
    
    strip_logic_pattern = re.compile(r"(    ' Desktop: Remove Hyperlink.*?On Error GoTo 0\n)", re.DOTALL)
    
    if strip_logic_pattern.search(content):
        # We need to indent the block and wrap it
        match = strip_logic_pattern.search(content).group(1)
        wrapped_logic = "    If Not isSetup Then\n" + match + "    End If\n"
        content = content.replace(match, wrapped_logic)
        print("Successfully wrapped stripping logic in ApplyDemoLockOnOpen")
    else:
        print("Could not find stripping logic to wrap.")

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
