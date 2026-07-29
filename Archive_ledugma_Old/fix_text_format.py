import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Increase shape size
    content = content.replace(', 220, 65)', ', 280, 80)')
    
    # Rebuild the string with 3 lines explicitly to prevent RTL jumbling
    line1 = ' & '.join(f'ChrW({ord(c)})' for c in "הקש כאן לפתיחת הוראות התקנה.")
    line2_part1 = ' & '.join(f'ChrW({ord(c)})' for c in "ניתן לשלוח ")
    line2_part2 = '"WhatsApp "'
    line2_part3 = ' & '.join(f'ChrW({ord(c)})' for c in "לטלפון ")
    line2_part4 = '"054-6677396"'
    line3 = ' & '.join(f'ChrW({ord(c)})' for c in "לקבלת עזרה בהתקנה.")
    
    vba_str = f'{line1} & vbCrLf & _\n{line2_part1} & {line2_part2} & {line2_part3} & {line2_part4} & vbCrLf & _\n{line3}'
    
    # We replace the text assignment. We need to catch the current text assignment block.
    # The current text assignment starts with `.Text = ChrW(1492)` and ends before `.Font.Size = 12`
    
    pattern = re.compile(r'\.Text = ChrW\(1492\).*?בהתקנה\."\s+', re.DOTALL)
    
    # Since I don't know exactly how my previous script generated the trailing characters,
    # let's just do a more robust regex. It starts with ".Text = ChrW(1492)" and ends with ".Font.Size"
    pattern = re.compile(r'\.Text = ChrW\(1492\).*?(?=\.Font\.Size =)', re.DOTALL)
    
    if pattern.search(content):
        new_content = pattern.sub(f'.Text = {vba_str}\n        ', content)
        with open(bas_file, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print("Successfully fixed text formatting!")
    else:
        # Fallback pattern
        print("Trying fallback regex...")
        pattern2 = re.compile(r'\.Text = .*?(?=\.Font\.Size =)', re.DOTALL)
        if pattern2.search(content):
            new_content = pattern2.sub(f'.Text = {vba_str}\n        ', content)
            with open(bas_file, 'w', encoding='utf-8') as f:
                f.write(new_content)
            print("Successfully fixed text formatting (fallback)!")
        else:
            print("Could not find the text block to replace.")

if __name__ == "__main__":
    main()
