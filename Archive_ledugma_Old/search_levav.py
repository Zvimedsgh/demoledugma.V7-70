import re

files_to_check = [
    r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas',
    r'C:\Users\Zvi\.gemini\antigravity\brain\d87047a2-5e99-4d1c-9993-702c3c3db5ae\Operation_Instructions.md'
]

for file_path in files_to_check:
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        for i, line in enumerate(lines):
            if re.search(r'LEVAV|לבב', line, re.IGNORECASE):
                print(f"{os.path.basename(file_path)}:{i+1}: {line.strip()}")
    except Exception as e:
        import os
        print(f"Error reading {file_path}: {e}")
