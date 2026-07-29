with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if 'BUILD_SUMMARY' in line or '2740' in line:
        print(f"Text Line {i+1}: {line.strip()}")
