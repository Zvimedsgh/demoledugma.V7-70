with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.254.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()
for i, line in enumerate(lines):
    if 'wsMain.Range("D19:K20").ClearContents' in line:
        print(f'Line {i}: {line.strip()}')
