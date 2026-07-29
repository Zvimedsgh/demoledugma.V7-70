text = """ChrW(1500) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & ChrW(32) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(32) & ChrW(1511) & ChrW(1500) & ChrW(1511) & ChrW(32) & ChrW(1499) & ChrW(1488) & ChrW(1503) & ChrW(32) & ChrW(1511) & ChrW(1500) & ChrW(1497) & ChrW(1511) & ChrW(32) & ChrW(1497) & ChrW(1502) & ChrW(1504) & ChrW(1497) & ChrW(32) & ChrW(1508) & vbCrLf & "054-6677396" & vbCrLf & ChrW(1513) & ChrW(1500) & ChrW(1495) & ChrW(32) & ChrW(1493) & ChrW(1493) & ChrW(1488) & ChrW(1496) & ChrW(1505) & ChrW(1488) & ChrW(1508) & ChrW(32) & ChrW(1500) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492)"""

res = ''
for part in text.split('&'):
    part = part.strip()
    if part.startswith('ChrW('):
        val = int(part[5:-1])
        res += chr(val)
    elif part == 'vbCrLf':
        res += '\n'
    elif part.startswith('"'):
        res += part.strip('"')

with open('decoded.txt', 'w', encoding='utf-8') as f:
    f.write(res)
