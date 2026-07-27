import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.30.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Change matachMsg to a single line without vbCrLf
    old_matach = 'matachMsg = ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1497) & ChrW(1501) & " " & _\n        ChrW(1489) & ChrW(1502) & ChrW(1496) & ChrW(34) & ChrW(1495) & " " & _\n        ChrW(1502) & ChrW(1514) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & vbCrLf & _\n        ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & _\n        ChrW(1492) & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & _\n        ChrW(1492) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1490) & vbCrLf & _\n        ChrW(1489) & ChrW(1504) & ChrW(1511) & " " & _\n        ChrW(1497) & ChrW(1513) & ChrW(1512) & ChrW(1488) & ChrW(1500) & " " & _\n        ChrW(1500) & "-15 " & _\n        ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & " " & _\n        ChrW(1492) & ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)'
    new_matach = 'matachMsg = ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1497) & ChrW(1501) & " " & ChrW(1489) & ChrW(1502) & ChrW(1496) & ChrW(34) & ChrW(1495) & " " & ChrW(1502) & ChrW(1514) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & ChrW(1492) & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & ChrW(1492) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1489) & ChrW(1504) & ChrW(1511) & " " & ChrW(1497) & ChrW(1513) & ChrW(1512) & ChrW(1488) & ChrW(1500) & " " & ChrW(1500) & "-15 " & ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & " " & ChrW(1492) & ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)'
    content = content.replace(old_matach, new_matach)

    # 2. Change range to J5:L7, UnMerge before merging, and set font to 9
    old_merge = 'wsMain.Range("J5:K7").Merge\n    wsMain.Range("J5").Value = matachMsg\n    wsMain.Range("J5").WrapText = True\n    wsMain.Range("J5").Font.Size = 6\n    wsMain.Range("J5").ShrinkToFit = False\n    wsMain.Range("J5").HorizontalAlignment = xlCenter\n    wsMain.Range("J5").VerticalAlignment = xlCenter\n    wsMain.Range("J5:K7").Interior.Color = RGB(230, 245, 255)'
    new_merge = 'wsMain.Range("J5:L7").UnMerge\n    wsMain.Range("J5:L7").Merge\n    wsMain.Range("J5").Value = matachMsg\n    wsMain.Range("J5").WrapText = True\n    wsMain.Range("J5").Font.Size = 9\n    wsMain.Range("J5").ShrinkToFit = False\n    wsMain.Range("J5").HorizontalAlignment = xlCenter\n    wsMain.Range("J5").VerticalAlignment = xlCenter\n    wsMain.Range("J5:L7").Interior.Color = RGB(230, 245, 255)'
    content = content.replace(old_merge, new_merge)

    # 3. Change borders logic to use J5:L7
    old_borders = 'For Each cellBorder In wsMain.Range("J5:K7")'
    new_borders = 'For Each cellBorder In wsMain.Range("J5:L7")'
    content = content.replace(old_borders, new_borders)

    # Bump version
    content = content.replace('Attribute VB_Name = "Goren_Claude1.30"', 'Attribute VB_Name = "Goren_Claude1.31"')
    content = content.replace("' VERSION: V1.30", "' VERSION: V1.31")
    content = content.replace('Private Const APP_VERSION As String = "1.30"', 'Private Const APP_VERSION As String = "1.31"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
