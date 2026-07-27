import os

folder = r'c:\LEVAV PROJECT\SOURCE'
for filename in os.listdir(folder):
    if filename.endswith('.bas') or filename.endswith('.cls'):
        filepath = os.path.join(folder, filename)
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for i, line in enumerate(f):
                if "שגיאה במקרו" in line or chr(1505) + chr(1504) + chr(1503) in line: # "סנן"
                    print(f"{filename} [{i+1}] {line.strip()}")

