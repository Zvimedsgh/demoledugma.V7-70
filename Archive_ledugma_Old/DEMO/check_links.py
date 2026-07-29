import fitz

doc = fitz.open(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = doc[0]
links = page.get_links()
for link in links:
    print("Found link:", link)
