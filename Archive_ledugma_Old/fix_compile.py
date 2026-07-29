def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    if 'Public Sub ResetHomeDefaults()' not in content:
        macro = """
Public Sub ResetHomeDefaults()
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    Application.EnableEvents = False
    
    ' G5
    wsMain.Range("G5").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    ' G6
    wsMain.Range("G6").Value = ""
    wsMain.Range("G6").Interior.Color = RGB(220, 240, 220)
    On Error Resume Next
    wsMain.Range("G6").Validation.Delete
    On Error GoTo 0
    wsMain.Range("G7").Value = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)
    ' G9
    wsMain.Range("G9").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)
    ' G10
    On Error Resume Next
    wsMain.Range("G10").Validation.Delete
    On Error GoTo 0
    wsMain.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)
    wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)
    
    Application.EnableEvents = True
    
    wsMain.Range("A1").Select
End Sub
"""
        with open(bas_file, 'a', encoding='utf-8') as f:
            f.write("\n" + macro)
        print("Fixed Compile Error!")
    else:
        print("Macro already exists.")

if __name__ == "__main__":
    main()
