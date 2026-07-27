import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.33.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Rename ApplyUserPermissions to CheckUserPermissions
    # Replace in comments
    content = content.replace("ApplyUserPermissions now disables buttons", "CheckUserPermissions now disables buttons")
    content = content.replace("APPLY USER PERMISSIONS: Show/hide sheets", "CHECK USER PERMISSIONS: Show/hide sheets")
    
    # Replace the sub name
    content = content.replace("Public Sub ApplyUserPermissions()", "Public Sub CheckUserPermissions()")

    # Write back
    with open("c:\\LEVAV PROJECT\\SOURCE\\tweak_layout_14.py.out", 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
