import fitz

doc = fitz.open(r'C:\ledugma\DEMO\Ad_190x60_PressQuality_v35.pdf')
page = doc[0]

# Search for the word "לפרטים"
text_instances = page.search_for("לפרטים")

url = "https://zgoren-my.sharepoint.com/personal/zvi_zgoren_com_il/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fzvi_zgoren_com_il%2FDocuments%2F%D7%9C%D7%93%D7%95%D7%92%D7%9E%D7%90%20%D7%A1%D7%95%D7%9B%D7%A0%D7%95%D7%AA%2FDemo_Reports_Syatem_V7.70.xlsm&parent=%2Fpersonal%2Fzvi_zgoren_com_il%2FDocuments%2F%D7%9C%D7%93%D7%95%D7%92%D7%9E%D7%90%20%D7%A1%D7%95%D7%9B%D7%A0%D7%95%D7%AA"

for inst in text_instances:
    # Expand the rectangle slightly to cover the whole button
    rect = inst
    rect.x0 -= 10
    rect.y0 -= 10
    rect.x1 += 30
    rect.y1 += 10
    
    # Add hyperlink
    link = {"kind": fitz.LINK_URI, "uri": url, "from": rect}
    page.insert_link(link)

doc.save(r'C:\ledugma\DEMO\Ad_190x60_PressQuality_v36.pdf')
print("Saved v36")
