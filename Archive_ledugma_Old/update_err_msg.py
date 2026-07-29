with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# We will replace the userMsg assignment in BuildPresentation
new_msg = """        Else
            userMsg = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & _
                      ChrW(1489) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & _
                      ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & vbCrLf & _
                      "Line: " & errLine & vbCrLf & "Error: " & errNum & " - " & errDesc
        End If"""

text = re.sub(r'Else\s+userMsg = ChrW\(1513\).*?errDesc\s+End If', new_msg, text, flags=re.DOTALL | re.IGNORECASE)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated error message!")
