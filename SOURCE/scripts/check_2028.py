import os
import sys

if os.path.exists(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas'):
    with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas', 'r', encoding='utf-8') as f:
        content = f.read()
        print("V2.028 exists! Length:", len(content))
        # Find SetupSettingsMenu
        import re
        m = re.search(r'Public Sub SetupSettingsMenu\(\).*?End Sub', content, re.DOTALL)
        if m:
            print("SetupSettingsMenu in V2.028:")
            for line in m.group(0).split('\n'):
                if 'OnAction' in line:
                    print(line.strip())
        else:
            print("SetupSettingsMenu not found in V2.028")
else:
    print("V2.028 does not exist")
