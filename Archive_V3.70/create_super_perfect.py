with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.59.bas', 'r', encoding='utf-8') as f:
    v359 = f.read()

start = v359.find('\' Final Summary Slide')
end = v359.find('Dim pg As Long', start)

final_slide = v359[start:end]

# Remove the coffee button and coffee slide from final_slide
coffee_btn_start = final_slide.find('\' Button 3: Coffee')
coffee_btn_end = final_slide.find('\' ===============================\n', final_slide.find('btnBack.ActionSettings(1).SoundEffect.Name = "Click"')) + 35

final_slide_no_coffee = final_slide[:coffee_btn_start] + final_slide[coffee_btn_end:]

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Final_System.bas', 'r', encoding='utf-8') as f:
    text = f.read()

text_start = text.find('\' Final Summary Slide')
text_end = text.find('Dim pg As Long', text_start)

new_text = text[:text_start] + final_slide_no_coffee + text[text_end:]

new_text = new_text.replace('Attribute VB_Name = "Goren_Claude_Final_System"', 'Attribute VB_Name = "Goren_Claude_Super_Perfect"')

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Super_Perfect.bas', 'w', encoding='utf-8') as f:
    f.write(new_text)

print('Super Perfect Slide created')
