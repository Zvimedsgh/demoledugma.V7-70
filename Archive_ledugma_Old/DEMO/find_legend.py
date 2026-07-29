import fitz

doc = fitz.open(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = doc[12]

words = page.get_text("words")
for w in words:
    # w is (x0, y0, x1, y1, "word", block_no, line_no, word_no)
    text = w[4]
    if "2019" in text or "2020" in text or "19" in text or "20" in text:
        print(w)
