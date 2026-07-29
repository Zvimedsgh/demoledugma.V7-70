import re

def patch_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Update the version string in the error handler
    content = content.replace('Error in V2.079!', 'Error in V2.224!')

    # 2. Fix the ApplyDemoLockOnOpen ERR_HANDLER issue
    # It currently ends with On Error GoTo ERR_HANDLER, which causes a compile error if it runs
    # We will find ApplyDemoLockOnOpen and replace the last On Error GoTo ERR_HANDLER with On Error GoTo 0
    # Wait, the injected shape ends with:
    #     wsMain.Hyperlinks.Add Anchor:=shpInstall, Address:="..."
    #     On Error GoTo ERR_HANDLER
    # We need to change the one specifically in ApplyDemoLockOnOpen.
    
    parts = content.split('Public Sub ApplyDemoLockOnOpen()')
    if len(parts) == 2:
        part2 = parts[1]
        # Replace the first occurrence of "On Error GoTo ERR_HANDLER" in ApplyDemoLockOnOpen
        # Wait, there might be one from the shape injection. Let's just replace the exact line.
        old_shape_err = 'wsMain.Hyperlinks.Add Anchor:=shpInstall, Address:="https://gorentec-my.sharepoint.com/:b:/g/personal/zvi_gorentech_co_il/IQBt2Ms0oWLySpRBl-6OY68MAXnBN2jyzpD3y43ZHmIUBwU?e=y10XmN"\n    On Error GoTo ERR_HANDLER'
        new_shape_err = 'wsMain.Hyperlinks.Add Anchor:=shpInstall, Address:="https://gorentec-my.sharepoint.com/:b:/g/personal/zvi_gorentech_co_il/IQBt2Ms0oWLySpRBl-6OY68MAXnBN2jyzpD3y43ZHmIUBwU?e=y10XmN"\n    On Error GoTo 0'
        
        if old_shape_err in part2:
            part2 = part2.replace(old_shape_err, new_shape_err, 1)
            content = parts[0] + 'Public Sub ApplyDemoLockOnOpen()' + part2
            print("Patched ApplyDemoLockOnOpen.")
        else:
            print("Warning: Could not find shape error handler in ApplyDemoLockOnOpen.")
    
    # 3. Comment out 4050 and 4110 to prevent 1004 during shutdown
    content = content.replace('4050 ThisWorkbook.Names("rngFILES_FOLDER").Delete', '4050 \'ThisWorkbook.Names("rngFILES_FOLDER").Delete')
    content = content.replace('4110 ThisWorkbook.Names("rngREPORTS_FOLDER").Delete', '4110 \'ThisWorkbook.Names("rngREPORTS_FOLDER").Delete')

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Patching complete.")

if __name__ == '__main__':
    patch_file(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.224.bas')
