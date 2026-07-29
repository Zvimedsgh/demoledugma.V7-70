with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
for i, line in enumerate(lines):
    if 'Sub SafeOpenWorkbook' in line:
        in_sub = True
    if in_sub:
        print(line.strip())
        if 'End Sub' in line:
            break
