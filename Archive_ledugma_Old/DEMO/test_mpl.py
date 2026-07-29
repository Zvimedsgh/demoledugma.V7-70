import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

fig = plt.figure(figsize=(7.48, 2.36))
ax = fig.add_axes([0, 0, 1, 1])
ax.axis('off')

text_str = "!דחא קילקבו האלמ תופיקשב - םיכותיחה לכ ,םינותנה לכ :תמדקתמ תוחוד תכרעמ"
ax.text(0.5, 0.08, text_str, fontsize=16, ha='center', va='center', fontfamily='Arial')

fig.savefig(r"C:\ledugma\DEMO\text_base.pdf", format='pdf')
print("PDF saved.")
