import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.26.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Clear old credit texts properly in SetupMainSheet
    old_credit_setup = '\' ---- Add Credit and Version to C23 ----\n    wsMain.Range("A23").Value ='
    new_credit_setup = '\' ---- Add Credit and Version to C23 ----\n    wsMain.Range("A20:Z30").ClearContents\n    wsMain.Range("A23").Value ='
    content = content.replace(old_credit_setup, new_credit_setup)

    # 2. Fix the Matach text to have explicit newlines and size 8
    old_matach = 'matachMsg = ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1497) & ChrW(1501) & " " & _\n        ChrW(1489) & ChrW(1502) & ChrW(1496) & ChrW(34) & ChrW(1495) & " " & _\n        ChrW(1502) & ChrW(1514) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & _\n        ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & _\n        ChrW(1492) & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & _\n        ChrW(1492) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & _\n        ChrW(1489) & ChrW(1504) & ChrW(1511) & " " & _\n        ChrW(1497) & ChrW(1513) & ChrW(1512) & ChrW(1488) & ChrW(1500) & " " & _\n        ChrW(1500) & "-15 " & _\n        ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & " " & _\n        ChrW(1492) & ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)'
    new_matach = 'matachMsg = ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1497) & ChrW(1501) & " " & _\n        ChrW(1489) & ChrW(1502) & ChrW(1496) & ChrW(34) & ChrW(1495) & " " & _\n        ChrW(1502) & ChrW(1514) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & vbCrLf & _\n        ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & _\n        ChrW(1492) & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & _\n        ChrW(1492) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1490) & vbCrLf & _\n        ChrW(1489) & ChrW(1504) & ChrW(1511) & " " & _\n        ChrW(1497) & ChrW(1513) & ChrW(1512) & ChrW(1488) & ChrW(1500) & " " & _\n        ChrW(1500) & "-15 " & _\n        ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & " " & _\n        ChrW(1492) & ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)'
    content = content.replace(old_matach, new_matach)

    # Change font size to 8 and remove ShrinkToFit
    content = content.replace('wsMain.Range("J5").WrapText = True\n    wsMain.Range("J5").Font.Size = 7\n    wsMain.Range("J5").ShrinkToFit = True', 'wsMain.Range("J5").WrapText = True\n    wsMain.Range("J5").Font.Size = 8\n    wsMain.Range("J5").ShrinkToFit = False')

    # Bump version
    content = content.replace('Attribute VB_Name = "Goren_Claude1.26"', 'Attribute VB_Name = "Goren_Claude1.27"')
    content = content.replace("' VERSION: V1.26", "' VERSION: V1.27")
    content = content.replace('Private Const APP_VERSION As String = "1.26"', 'Private Const APP_VERSION As String = "1.27"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
