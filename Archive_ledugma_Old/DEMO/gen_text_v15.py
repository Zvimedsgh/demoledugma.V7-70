import matplotlib.pyplot as plt

fig = plt.figure(figsize=(7.48, 2.36))
fig.patch.set_alpha(0.0)

ax = fig.add_axes([0, 0, 1, 1])
ax.axis('off')

text1 = "מערכת דיווח לסוכנויות/סוכני ביטוח: כל הנתונים, כל החתכים - בשקיפות מלאה ובקליק אחד!"
text2 = "גמישות בהתאמה לצרכי הסוכן, תמיכה ושירות מקצועי."

ax.text(0.55, 0.12, text1[::-1], fontsize=11, ha='center', va='center', fontfamily='Arial')
ax.text(0.55, 0.05, text2[::-1], fontsize=11, ha='center', va='center', fontfamily='Arial')

button_text = "< ץחל םיטרפל"
# Halfway between 0.12 and 0.06 is 0.09
ax.text(0.09, 0.085, button_text, fontsize=12, ha='center', va='center', fontfamily='Arial', color='white', fontweight='bold',
        bbox=dict(boxstyle="round,pad=0.5", fc="darkorange", ec="none"))

fig.savefig(r"C:\ledugma\DEMO\text_base_trans_v15.pdf", format='pdf', transparent=True)
print("Transparent text v15 created.")
