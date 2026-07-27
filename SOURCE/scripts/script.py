import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.017.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = '''120         listName = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504) & ChrW(1497) & "," & ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & "," & ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497) & "," & ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497) & " " & ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497)'''

new_code = '''120         listName = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503) & "," & ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & "," & ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497) & "," & ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497)'''

if old_code not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code, new_code)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
