import os
import re

directory = r'c:\LEVAV PROJECT\SOURCE'
for filename in os.listdir(directory):
    if filename.endswith('.bas') or filename.endswith('.cls'):
        filepath = os.path.join(directory, filename)
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()
            if re.search(r'(Sub MANUAL\b|Function MANUAL\b)', text, re.IGNORECASE):
                print(f"Found in {filename}")
                match = re.search(r'(Sub MANUAL\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
                if match:
                    print(match.group(1))
