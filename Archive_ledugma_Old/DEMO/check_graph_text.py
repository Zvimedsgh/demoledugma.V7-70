import fitz

doc = fitz.open(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = doc[12] # Graph
words = page.get_text("words")
with open(r"C:\ledugma\DEMO\check_graph_text.txt", "w", encoding="utf-8") as f:
    for w in words:
        f.write(f"X:{w[0]:.0f}-{w[2]:.0f}, Y:{w[1]:.0f}-{w[3]:.0f} -> {w[4]}\n")
