with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.58.bas', 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V3_58"', 'Attribute VB_Name = "Goren_Claude_V3_59"')
content = content.replace('coffee.jpg', 'coffee.png')

old_pic_code = '''            Set picCoffee = sldCoffee.Shapes.AddShape(1, 230, 20, 500, 500)
            picCoffee.Fill.UserPicture coffeePath
            picCoffee.Line.Visible = 0'''

new_pic_code = '''            Set picCoffee = sldCoffee.Shapes.AddPicture(coffeePath, 0, -1, 230, 20, 500, 500)'''

content = content.replace(old_pic_code, new_pic_code)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.59.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print("V3.59 created")
