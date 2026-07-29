def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Replace both instances of OpenInstallWeb with OpenInstructions
    new_content = content.replace('shpInstall.OnAction = "OpenInstallWeb"', 'shpInstall.OnAction = "OpenInstructions"')
    
    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(new_content)
        
    print("Successfully restored OnAction to OpenInstructions!")

if __name__ == "__main__":
    main()
