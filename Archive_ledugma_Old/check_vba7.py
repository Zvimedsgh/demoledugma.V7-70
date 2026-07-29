with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if 'VBA7' in line or 'PtrSafe' in line or '#Else' in line or '#End If' in line:
        print(f"Line {i+1}: {line.strip()}")
