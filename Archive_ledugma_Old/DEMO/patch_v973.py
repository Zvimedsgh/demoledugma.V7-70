import datetime

def create_v973():
    with open(r'C:\ledugma\DEMO\V9.72_Final.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    # Update version
    code = code.replace("9.72", "9.73")
    code = code.replace("V9.72", "V9.73")

    # Change CONTROL_SHEET_NAME to "דף הבית"
    old_func = """Public Function CONTROL_SHEET_NAME() As String
    ' Returns "ראשי"
    CONTROL_SHEET_NAME = ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1497)
End Function"""

    new_func = """Public Function CONTROL_SHEET_NAME() As String
    ' Returns "דף הבית"
    CONTROL_SHEET_NAME = ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
End Function"""

    code = code.replace(old_func, new_func)

    with open(r'C:\ledugma\DEMO\V9.73_Final.txt', 'w', encoding='windows-1255') as f:
        f.write(code)
        
    print("V9.73_Final.txt created successfully.")

if __name__ == '__main__':
    create_v973()
