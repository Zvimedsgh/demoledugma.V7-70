import matplotlib.pyplot as plt

fig = plt.figure(figsize=(7.48, 2.36))
fig.patch.set_alpha(0.0)

ax = fig.add_axes([0, 0, 1, 1])
ax.axis('off')

text1 = "מערכת דוחות מתקדמת: כל הנתונים, כל החיתוכים - בשקיפות מלאה ובקליק אחד!"
text2 = "גמישות בהתאמה לצרכי הסוכן, תמיכה ושירות מקצועי."

ax.text(0.5, 0.12, text1[::-1], fontsize=13, ha='center', va='center', fontfamily='Arial')
ax.text(0.5, 0.05, text2[::-1], fontsize=13, ha='center', va='center', fontfamily='Arial')

fig.savefig(r"C:\ledugma\DEMO\text_base_trans_v10.pdf", format='pdf', transparent=True)
print("Transparent text v10 created.")
