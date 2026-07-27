import sys
import os

folder = r'c:\LEVAV PROJECT\SOURCE'
for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('.bas') or file.endswith('.vba'):
            filepath = os.path.join(root, file)
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                if "עושים" in content or "הסבר" in content or "כפתורים" in content or "כפתור 1" in content:
                    print(f"Match in {file}")
