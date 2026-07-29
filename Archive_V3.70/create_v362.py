with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.61.bas', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('Attribute VB_Name = "Goren_Claude_V3_61"', 'Attribute VB_Name = "Goren_Claude_V3_62"')

text = text.replace("' VERSION: V3.50\n' CHANGES IN V3.43:", "' VERSION: V3.62\n' CHANGES IN V3.62: Final robust version.\n' CHANGES IN V3.43:")

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.62.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print('V3.62 created')
