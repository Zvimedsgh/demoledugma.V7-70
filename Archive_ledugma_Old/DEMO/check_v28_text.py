import fitz

doc = fitz.open(r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v28.pdf")
page = doc[0]
words = page.get_text("words")
with open(r"C:\ledugma\DEMO\v28_text.txt", "w", encoding="utf-8") as f:
    for w in words:
        if w[0] > 250:
            f.write(f"X:{w[0]:.0f}-{w[2]:.0f}, Y:{w[1]:.0f}-{w[3]:.0f} -> {w[4]}\n")
