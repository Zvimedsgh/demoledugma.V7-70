import re
import datetime

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    out_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.223.bas'
    
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Fix the missing apostrophes
    fixes = [
        (" A1   = ", "' A1   = "),
        (" A57  = ", "' A57  = "),
        (" A173 = ", "' A173 = "),
        (" A181 = ", "' A181 = "),
        (" A196 = ", "' A196 = "),
        (" A221 = ", "' A221 = "),
        (" A240 = ", "' A240 = "),
        (" A249 = ", "' A249 = ")
    ]
    
    for old_s, new_s in fixes:
        content = content.replace("\n" + old_s, "\n" + new_s)
        content = content.replace("\r\n" + old_s, "\r\n" + new_s)

    # Update version header
    content = content.replace('VERSION: V2.222', 'VERSION: V2.223')
    content = content.replace('Attribute VB_Name = "Goren_Claude_V2_222"', 'Attribute VB_Name = "Goren_Claude_V2_223"')
    
    # Update timestamp
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    content = re.sub(r"' DATE: .*?\(LATEST FIXES\)", f"' DATE: {now_str} (LATEST FIXES)", content)

    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(content)
        
    print(f"Created {out_file} successfully.")

if __name__ == "__main__":
    main()
