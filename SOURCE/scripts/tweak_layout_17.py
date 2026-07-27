import re
import datetime

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.36.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # We need to insert `    SetupSectionNamedRanges` right before `    periodBaseRow = ThisWorkbook.Names("rngSection_PeriodLists").RefersToRange.Row`
    # Let's find that line
    target_line = r'      periodBaseRow = ThisWorkbook.Names\("rngSection_PeriodLists"\).RefersToRange.Row'
    if not re.search(target_line, content):
        print("Could not find periodBaseRow assignment!")
        return

    # Insert SetupSectionNamedRanges
    replacement = r"      ' Initialize Named Ranges first so we can reference them below\n      SetupSectionNamedRanges\n\n      periodBaseRow = ThisWorkbook.Names(\"rngSection_PeriodLists\").RefersToRange.Row"
    content = re.sub(target_line, replacement, content)

    # Replace VB_Name
    content = re.sub(r'Attribute VB_Name = "[^"]+"', 'Attribute VB_Name = "Goren_Claude1_37"', content)

    # Replace VERSION header
    content = re.sub(r"' VERSION: V\d+\.\d+", "' VERSION: V1.37", content)

    # Replace APP_VERSION constant
    content = re.sub(r'Private Const APP_VERSION As String = "\d+\.\d+"', 'Private Const APP_VERSION As String = "1.37"', content)

    # Replace DATE header
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    content = re.sub(r"' DATE: \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", f"' DATE: {current_time}", content)

    # Add changelog
    changelog = f"""' ============================================================================
' CHANGES IN 1.37:
'   - Fixed: SetupMainSheet crashing on first run (Error 91) due to missing rngSection_PeriodLists and rngSection_Messages Named Ranges. SetupSectionNamedRanges is now called at the beginning of SetupMainSheet."""

    content = content.replace("' ============================================================================\n' CHANGES IN 1.36", changelog + "\n' CHANGES IN 1.36")

    # Write back to V1.37
    with open(r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.37.bas", 'w', encoding='utf-8') as f:
        f.write(content)

    print("Created Goren_Claude_V1.37.bas")

if __name__ == "__main__":
    main()
