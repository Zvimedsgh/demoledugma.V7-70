import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.224.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Change version text
    content = content.replace('Error in V2.224!', 'Error in V2.225!')
    
    # We will completely remove the numbered lines 4020 to 4150 and replace them
    # with ultra-safe, unnumbered VBA code.
    old_block = """    4020 On Error Resume Next
    4030 existingPath = Trim$(CStr(ThisWorkbook.Names("rngFILES_FOLDER").RefersToRange.Value2))
    4040 If existingPath = "" Then
        4050 'ThisWorkbook.Names("rngFILES_FOLDER").Delete
        4060 Err.Clear
        4070 ThisWorkbook.Names.Add "rngFILES_FOLDER", ThisWorkbook.Worksheets(H_SET_PARAMS()).Range("B176")
    4080 End If
    4090 existingPath = Trim$(CStr(ThisWorkbook.Names("rngREPORTS_FOLDER").RefersToRange.Value2))
    4100 If existingPath = "" Then
        4110 'ThisWorkbook.Names("rngREPORTS_FOLDER").Delete
        4120 Err.Clear
        4130 ThisWorkbook.Names.Add "rngREPORTS_FOLDER", ThisWorkbook.Worksheets(H_SET_PARAMS()).Range("B177")
    4140 End If
    4150 Err.Clear"""

    new_block = """    On Error Resume Next
    Dim dummyErr As Long
    ThisWorkbook.Names.Add "rngFILES_FOLDER", ThisWorkbook.Worksheets(H_SET_PARAMS()).Range("B176")
    ThisWorkbook.Names.Add "rngREPORTS_FOLDER", ThisWorkbook.Worksheets(H_SET_PARAMS()).Range("B177")
    Err.Clear"""

    if old_block in content:
        content = content.replace(old_block, new_block)
        print("Successfully replaced block 4020-4150.")
    else:
        print("Could not find block 4020-4150!")
        
    out_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.225.bas'
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Created {out_file}")

if __name__ == '__main__':
    main()
