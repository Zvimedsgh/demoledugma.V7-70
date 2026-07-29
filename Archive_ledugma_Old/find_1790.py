with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i, line in enumerate(lines):
    if '1790    ThisWorkbook.Protect "Z961814r"' in line:
        print(f"Line {i+1}: {line.strip()}")
        for j in range(max(0, i-5), min(len(lines), i+5)):
            if 'Sub ' in lines[j] or 'Function ' in lines[j]:
                print(f"  {lines[j].strip()}")
