with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i, line in enumerate(lines):
    if line.strip().startswith('2740'):
        print(f"Found 2740 at line {i+1}: {line.strip()}")
        # context
        for j in range(max(0, i-4), min(len(lines), i+5)):
            print(f"  {j+1}: {lines[j].strip()}")
