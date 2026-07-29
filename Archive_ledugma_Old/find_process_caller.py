with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i, line in enumerate(lines):
    if 'ProcessSummaryData' in line and 'Sub ProcessSummaryData' not in line:
        print(f"Line {i+1}: {line.strip()}")
        # Context
        for j in range(max(0, i-5), i):
            if 'Sub ' in lines[j]:
                print(f"  Called from: {lines[j].strip()}")
