import os

found_files = []
for root, dirs, files in os.walk(r'C:\Users\Zvi'):
    if 'OneDrive' in root:
        for f in files:
            if f.lower() == 'manual.docx':
                found_files.append(os.path.join(root, f))

for f in found_files:
    print(f)
