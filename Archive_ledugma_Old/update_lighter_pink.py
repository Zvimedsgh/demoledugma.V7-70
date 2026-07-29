with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

# Only replace the fill color for shpInstall
text = text.replace("shpInstall.Fill.ForeColor.RGB = RGB(255, 153, 153)", "shpInstall.Fill.ForeColor.RGB = RGB(255, 200, 200)")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated shape fill to lighter pink successfully.")
