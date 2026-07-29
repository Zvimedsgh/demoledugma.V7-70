import fitz
doc = fitz.open(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = doc[12] # graph page
pix = page.get_pixmap(dpi=150)
pix.save(r"C:\ledugma\DEMO\page13.png")
