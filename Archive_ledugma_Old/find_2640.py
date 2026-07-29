with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i, line in enumerate(lines):
    if '2640    BuildComparisonSheet' in line:
        for j in range(max(0, i-10), i+20):
            print(f"{j+1}: {lines[j].strip()}")
        break
