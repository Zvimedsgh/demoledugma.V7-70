with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i, line in enumerate(lines):
    if '2740    BuildSummarySheet' in line:
        for j in range(i, i+15):
            print(f"{j+1}: {lines[j].strip()}")
        break
