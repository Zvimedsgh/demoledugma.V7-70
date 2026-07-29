import fitz
doc = fitz.open(r'C:\ledugma\DEMO\מצגת הנהלה 2020.pdf')
page13 = doc[12]
page14 = doc[13]
with open('text_out.txt', 'w', encoding='utf-8') as f:
    f.write('Page 13:\n' + page13.get_text('text') + '\n')
    f.write('Page 14:\n' + page14.get_text('text') + '\n')
