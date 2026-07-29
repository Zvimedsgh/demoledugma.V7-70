with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
print(f"Line 6640 context:")
for i in range(6630, 6645):
    print(lines[i].strip())

print(f"\nLine 6782 context:")
for i in range(6770, 6785):
    print(lines[i].strip())
