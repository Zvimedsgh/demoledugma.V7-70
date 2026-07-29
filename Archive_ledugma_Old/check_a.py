with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.258.bas', 'r', encoding='utf-8') as f:
    for line in f:
        if 'wsMain.Range("A' in line and '.Value' in line:
            print(line.strip())
