import re
from datetime import datetime

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Update the date and time
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # The file has: ' DATE: 2026-07-03
    content = re.sub(r"' DATE: \d{4}-\d{2}-\d{2}.*", f"' DATE: {now_str} (LATEST FIXES)", content)

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)
        
if __name__ == "__main__":
    main()
