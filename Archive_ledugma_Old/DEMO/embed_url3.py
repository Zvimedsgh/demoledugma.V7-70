import fitz

doc = fitz.open(r'C:\ledugma\DEMO\Ad_190x60_PressQuality_v35.pdf')
page = doc[0]

url = "https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQCXXoKtDl4LSLKW3Ox0cUWdAWFpeuCFAi73S63Qowledio?e=nbuMsU"

found = False
# Try different representations of the text
for text in ["לפרטים", "םיטרפל", "לחץ", "ץחל", "<", ">", "לפרטים לחץ <", "> ץחל םיטרפל"]:
    instances = page.search_for(text)
    if instances:
        for inst in instances:
            rect = inst
            rect.x0 -= 15
            rect.y0 -= 15
            rect.x1 += 35
            rect.y1 += 15
            page.insert_link({"kind": fitz.LINK_URI, "uri": url, "from": rect})
            found = True

# If text search fails completely because it's vector paths and not text, we just add a manual rectangle
if not found:
    # Button is on the bottom-left roughly (x: 20-150, y: 120-160 for a 538x170 doc)
    # The doc is 538x170 (190x60mm). Left is x=0, top is y=0.
    # The button was placed at x=20 to x=150, y=120 to y=160
    rect = fitz.Rect(10, 110, 170, 160)
    page.insert_link({"kind": fitz.LINK_URI, "uri": url, "from": rect})

doc.save(r'C:\ledugma\DEMO\Ad_190x60_PressQuality_v38.pdf')
print("Saved v38")
