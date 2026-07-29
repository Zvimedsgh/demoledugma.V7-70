import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # The new text for shpInstallMsg: ONLY the WhatsApp lines.
    line1_part1 = ' & '.join(f'ChrW({ord(c)})' for c in "ניתן לשלוח הודעת ")
    line1_part2 = '"WhatsApp "'
    line1_part3 = ' & '.join(f'ChrW({ord(c)})' for c in "לטלפון")
    line2 = '"054-6677396"'
    line3 = ' & '.join(f'ChrW({ord(c)})' for c in "לקבלת עזרה.")
    
    vba_str = f'{line1_part1} & {line1_part2} & {line1_part3} & vbCrLf & _\n{line2} & vbCrLf & _\n{line3}'
    
    # We replace the text inside With shpInstall.TextFrame2.TextRange
    pattern_install = re.compile(r'(    With shpInstall\.TextFrame2\.TextRange\n        \.Text = ).*?(\n        \.Font\.Size = 12)', re.DOTALL)
    if pattern_install.search(content):
        content = pattern_install.sub(r'\1' + vba_str + r'\2', content)
        print("Updated shpInstallMsg text.")
    else:
        print("Could not find shpInstallMsg text block.")

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
