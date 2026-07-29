with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.258.bas', 'r', encoding='utf-8') as f:
    for line in f:
        if any(x in line for x in ['J2', 'K2', 'J3', 'K3', 'J4', 'K4', 'J5', 'K5', 'K7', 'L1']):
            print(line.strip())
