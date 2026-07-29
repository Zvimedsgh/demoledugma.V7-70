with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(4168, 0, -1):
    if lines[i].startswith('Sub '):
        print(lines[i].strip())
        break
