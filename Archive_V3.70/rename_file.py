with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Orit_Final.bas', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('Attribute VB_Name = "Goren_Claude_Orit_Final"', 'Attribute VB_Name = "Goren_claude_Orit_final_V3_70"')

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_claude_Orit_final_V3.70.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print('Created new file')
