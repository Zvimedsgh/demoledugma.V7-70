import re

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Orit_Final.bas', 'r', encoding='utf-8') as f:
    text = f.read()

# Replace APP_VERSION value
text = re.sub(r'APP_VERSION As String = ".*?"', 'APP_VERSION As String = "8.00"', text)

# Replace the text assignment to A19
old_text = 'wsMain.Range("A19").Value = ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " v" & APP_VERSION'
new_text = 'wsMain.Range("A19").Value = ChrW(1490) & ChrW(1497) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " V" & APP_VERSION'

text = text.replace(old_text, new_text)

# Do the same for creditText in the final slide/reports (if it exists)
old_credit = 'creditText = ChrW(1504) & ChrW(1489) & ChrW(1504) & ChrW(1492) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1497) & _'
new_credit = 'creditText = ChrW(1490) & ChrW(1497) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " V" & APP_VERSION'
# The block actually has a few lines
# 8729: creditText = ChrW(1504) & ChrW(1489) & ChrW(1504) & ChrW(1492) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1497) & _
# 8730:     ChrW(1491) & ChrW(1497) & " " & ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " v" & APP_VERSION
# Let's just regex replace creditText assignment up to APP_VERSION
text = re.sub(r'creditText = ChrW\(1504\).*?APP_VERSION', 'creditText = ChrW(1490) & ChrW(1497) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " V" & APP_VERSION', text, flags=re.DOTALL)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Orit_Final.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print('Updated versions to V8.00 and changed text to גירסה')
