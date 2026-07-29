with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i, line in enumerate(lines):
    if ('.Left =' in line or '.Top =' in line) and ('shp' in line or 'Shape' in line or 'Msg' in line or 'btn' in line):
        print(f"Line {i+1}: {line.strip()}")
