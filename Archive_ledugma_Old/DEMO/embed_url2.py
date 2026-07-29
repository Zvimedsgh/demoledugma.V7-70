import fitz

doc = fitz.open(r'C:\ledugma\DEMO\Ad_190x60_PressQuality_v35.pdf')
page = doc[0]

text_instances = page.search_for("לפרטים")

url = "https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQCXXoKtDl4LSLKW3Ox0cUWdAWFpeuCFAi73S63Qowledio?e=nbuMsU"

for inst in text_instances:
    rect = inst
    rect.x0 -= 10
    rect.y0 -= 10
    rect.x1 += 30
    rect.y1 += 10
    
    link = {"kind": fitz.LINK_URI, "uri": url, "from": rect}
    page.insert_link(link)

doc.save(r'C:\ledugma\DEMO\Ad_190x60_PressQuality_v37.pdf')
print("Saved v37")
