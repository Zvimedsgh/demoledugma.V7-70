with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
for i, line in enumerate(lines):
    if 'BuildSummarySheet' in line and 'Sub BuildSummarySheet' not in line:
        print(f"Line {i+1}: {line.strip()}")
        # Check context
        for j in range(max(0, i-5), i):
            print(f"  {j+1}: {lines[j].strip()}")
