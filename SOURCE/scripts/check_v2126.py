import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.126.bas'

import os
if os.path.exists(filepath):
    print("V2.126 exists.")
    with open(filepath, 'r', encoding='utf-8') as f:
        print(f"File size: {len(f.read())}")
else:
    print("V2.126 does not exist.")

