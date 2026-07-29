import pypdf

reader = pypdf.PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = reader.pages[6]
print("Slide width:", page.mediabox.width)
print("Slide height:", page.mediabox.height)
