with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

replacement = """    With wsMain.Range("F14:K15")
        .Merge
        .Value = ChrW(1512) & ChrW(1488) & ChrW(1492) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & "," & " " & ChrW(1500) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & "W" & "h" & "a" & "t" & "s" & "A" & "p" & "p" & " " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & " " & "0" & "5" & "4" & "-" & "6" & "6" & "7" & "7" & "3" & "9" & "6" """

import re
# Find the exact block starting with With wsMain.Range("D19:K20")
pattern = r'With wsMain\.Range\("D19:K20"\).*?": 054-6677396"'
text = re.sub(pattern, replacement.strip(), text, flags=re.DOTALL)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated successfully via python.")
