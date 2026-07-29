with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i, line in enumerate(lines):
    if 'Protect' in line and 'ThisWorkbook' in line and 'Z961814r' not in line:
        if not line.strip().startswith("'"):
            print(f"Line {i+1}: {line.strip()}")
