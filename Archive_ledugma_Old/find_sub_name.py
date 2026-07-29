with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i in range(1744, 0, -1):
    if 'Sub ' in lines[i] or 'Function ' in lines[i]:
        print(f"Line {i+1}: {lines[i].strip()}")
        break
