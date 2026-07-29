with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3845, 3855):
    print(f"{i+1}: {lines[i].strip()}")
