with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Final_System.bas', 'r', encoding='utf-8') as f:
    text = f.read()

start = text.find('.Text = ChrW(1492)')
end = text.find('.ParagraphFormat.Alignment = 2', start)

clean_text = '''            .Text = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & _
                    ChrW(1492) & ChrW(1493) & ChrW(1508) & ChrW(1511) & ChrW(1492) & " " & _
                    ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!"
'''

text = text[:start] + clean_text + text[end:]

# Update VB_Name
text = text.replace('Attribute VB_Name = "Goren_Claude_Final_System"', 'Attribute VB_Name = "Goren_Claude_Final_NoCoffee"')
text = text.replace('VERSION: Final_System', 'VERSION: Final_NoCoffee')

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Final_NoCoffee.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print('Created NoCoffee')
