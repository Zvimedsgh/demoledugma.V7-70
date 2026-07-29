with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i in range(max(0, 8680), min(len(lines), 8710)):
    print(f"{i+1}: {lines[i].strip()}")
