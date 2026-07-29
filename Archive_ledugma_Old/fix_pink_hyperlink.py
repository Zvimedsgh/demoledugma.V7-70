import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # The wrong URL (Operation Manual PDF)
    wrong_url = 'https://gorentec-my.sharepoint.com/:b:/g/personal/zvi_gorentech_co_il/IQBt2Ms0oWLySpRBl-6OY68MAXnBN2jyzpD3y43ZHmIUBwU?e=y10XmN'
    
    # The correct URL (Installation Instructions Excel)
    correct_url = 'https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=Z8iVEV'

    # I will replace ONLY where it is assigned to the Anchor (which is what I injected)
    # to avoid breaking the actual OpenManual macro
    target_string_wrong = f'Anchor:=shpInstall, Address:="{wrong_url}"'
    target_string_correct = f'Anchor:=shpInstall, Address:="{correct_url}"'

    if target_string_wrong in content:
        content = content.replace(target_string_wrong, target_string_correct)
        print("Replaced wrong URL with correct URL in shpInstallMsg hyperlinks.")
    else:
        print("Could not find the wrong URL in hyperlink assignments.")

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
