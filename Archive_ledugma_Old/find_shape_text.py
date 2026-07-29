with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i in range(len(lines)):
    if 'TextRange.Text' in lines[i] and ('הוראות' in lines[i] or 'ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514)' in lines[i]):
        print(f"Line {i+1}: {lines[i].strip()}")
        print(f"Shape: {lines[i-1].strip()}")
