with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

target1 = 'SOURCE_FOLDER = "C:\\" & ChrW(1508) & ChrW(1512) & ChrW(1493) & ChrW(1497) & ChrW(1511) & ChrW(1496) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489) & "\\SOURCE\\"'
text = text.replace(target1, 'SOURCE_FOLDER = ThisWorkbook.Path & "\\SOURCE\\"')

target2 = 'wsParams.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = "C:\\LEVAV PROJECT\\BACKUPS"'
text = text.replace(target2, 'wsParams.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = ThisWorkbook.Path & "\\BACKUPS"')

target3 = 'Open "c:\\LEVAV PROJECT\\SOURCE\\debug_log.txt" For Append As #ff'
text = text.replace(target3, 'Open ThisWorkbook.Path & "\\debug_log.txt" For Append As #ff')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print("Hardcoded paths removed.")
