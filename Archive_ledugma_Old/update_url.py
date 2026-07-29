import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Update OpenInstallWeb to use the actual URL
    old_url_code = 'ActiveWorkbook.FollowHyperlink "https://gorentec-my.sharepoint.com/"'
    new_url_code = 'ActiveWorkbook.FollowHyperlink "https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=Z8iVEV"'
    
    if old_url_code in content:
        content = content.replace(old_url_code, new_url_code)
    
    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Success updating URL!")

if __name__ == "__main__":
    main()
