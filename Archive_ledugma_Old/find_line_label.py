with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

for line in text.split('\n'):
    if line.strip().startswith('4850 '):
        print(line.strip())
