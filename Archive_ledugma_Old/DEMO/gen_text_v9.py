import matplotlib.pyplot as plt

fig = plt.figure(figsize=(7.48, 2.36))
fig.patch.set_alpha(0.0)  # Transparent background
ax = fig.add_axes([0, 0, 1, 1])
ax.axis('off')

# The text to render
original_text = "מערכת דוחות מתקדמת: כל הנתונים, כל החיתוכים - בשקיפות מלאה ובקליק אחד! גמישות בהתאמה לצרכי הסוכן, תמיכה ושירות מקצועי."

# Reverse the text for LTR rendering
reversed_text = original_text[::-1]

# Font size reduced from 16 to 12
ax.text(0.5, 0.08, reversed_text, fontsize=12, ha='center', va='center', fontfamily='Arial')

fig.savefig(r"C:\ledugma\DEMO\text_base_trans_v9.pdf", format='pdf', transparent=True)
print("Transparent text PDF saved.")
