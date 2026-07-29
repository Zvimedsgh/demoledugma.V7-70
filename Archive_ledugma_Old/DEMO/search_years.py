import fitz

input_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf"
doc = fitz.open(input_pdf)

hits_2019 = 0
hits_2020 = 0

for page in doc:
    rects_19 = page.search_for("2019")
    rects_20 = page.search_for("2020")
    if rects_19:
        print(f"Page {page.number} has {len(rects_19)} hits for 2019")
        hits_2019 += len(rects_19)
    if rects_20:
        print(f"Page {page.number} has {len(rects_20)} hits for 2020")
        hits_2020 += len(rects_20)

print(f"Total 2019 hits: {hits_2019}")
print(f"Total 2020 hits: {hits_2020}")
