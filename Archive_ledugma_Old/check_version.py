with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

for line in text.split('\n'):
    if '2.213' in line or '2.222' in line or 'Version' in line:
        print(line.strip())
