import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Update shpInstallMsg text
    # The text assignment starts with .Text = ChrW(1492) and ends with ChrW(46) before .Font.Size = 12
    # We will find the exact With shpInstall.TextFrame2.TextRange block.
    
    line1 = ' & '.join(f'ChrW({ord(c)})' for c in "הקש כאן לפתיחת הוראות התקנה")
    line2_part1 = ' & '.join(f'ChrW({ord(c)})' for c in "ניתן לשלוח הודעת ")
    line2_part2 = '"WhatsApp "'
    line2_part3 = ' & '.join(f'ChrW({ord(c)})' for c in "לטלפון")
    line3 = '"054-6677396"'
    line4 = ' & '.join(f'ChrW({ord(c)})' for c in "לקבלת עזרה.")
    
    vba_str = f'{line1} & vbCrLf & _\n{line2_part1} & {line2_part2} & {line2_part3} & vbCrLf & _\n{line3} & vbCrLf & _\n{line4}'
    
    # We replace the text inside With shpInstall.TextFrame2.TextRange
    pattern_install = re.compile(r'(    With shpInstall\.TextFrame2\.TextRange\n        \.Text = ).*?(\n        \.Font\.Size = 12)', re.DOTALL)
    if pattern_install.search(content):
        content = pattern_install.sub(r'\1' + vba_str + r'\2', content)
        print("Updated shpInstallMsg text.")
    else:
        print("Could not find shpInstallMsg text block.")

    # 2. Remove shpDismissMsg creation
    # Pattern to remove:
    #     Dim shpDismiss As Shape
    #     ...
    #     shpDismiss.OnAction = "DismissInstallMsg"
    pattern_dismiss = re.compile(r'    Dim shpDismiss As Shape\n.*?shpDismiss\.OnAction = "DismissInstallMsg"\n', re.DOTALL)
    if pattern_dismiss.search(content):
        content = pattern_dismiss.sub('', content)
        print("Removed shpDismissMsg creation.")
    else:
        print("Could not find shpDismissMsg creation block.")

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
