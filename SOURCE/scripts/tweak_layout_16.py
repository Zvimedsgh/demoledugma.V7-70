import re
import datetime

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.35.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # Replace VB_Name
    content = re.sub(r'Attribute VB_Name = "[^"]+"', 'Attribute VB_Name = "Goren_Claude1_36"', content)

    # Replace VERSION header
    content = re.sub(r"' VERSION: V\d+\.\d+", "' VERSION: V1.36", content)

    # Replace APP_VERSION constant
    content = re.sub(r'Private Const APP_VERSION As String = "\d+\.\d+"', 'Private Const APP_VERSION As String = "1.36"', content)

    # Replace DATE header
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    content = re.sub(r"' DATE: \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", f"' DATE: {current_time}", content)

    # Add changelog
    changelog = f"""' ============================================================================
' CHANGES IN 1.36:
'   - Fixed: SetupSettingsMenu now freezes top 4 rows and places navigation buttons horizontally as a floating toolbar.
'   - Fixed: SetupSectionNamedRanges dynamically offsets ranges based on toolbar presence.
'   - Fixed: Renamed ApplyUserPermissions back to CheckUserPermissions to fix Workbook_Open macro error.
'   - Updated: Version numbering, VB_Name, and timestamp headers in code."""

    content = content.replace("' ============================================================================\n' CHANGES IN 9.28", changelog + "\n' CHANGES IN 9.28")

    # Write back
    with open(r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.36.bas", 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
