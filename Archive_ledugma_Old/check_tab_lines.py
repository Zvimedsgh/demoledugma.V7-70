with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(9010, 9025):
    print(f"{i+1}: {lines[i].strip()}")

print("------")

for i in range(9125, 9135):
    print(f"{i+1}: {lines[i].strip()}")
