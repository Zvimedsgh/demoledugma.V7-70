import win32com.client
import os

try:
    ppApp = win32com.client.Dispatch('PowerPoint.Application')
    img_path = os.path.abspath('test_img.jpg')
    with open(img_path, 'wb') as f:
        f.write(b'\xff\xd8\xff\xe0\x00\x10JFIF\x00\x01\x01\x01\x00H\x00H\x00\x00\xff\xdb\x00C\x00\x08\x06\x06\x07\x06\x05\x08\x07\x07\x07\t\t\x08\n\x0c\x14\r\x0c\x0b\x0b\x0c\x19\x12\x13\x0f\x14\x1d\x1a\x1f\x1e\x1d\x1a\x1c\x1c $.\' \",#\x1c\x1c(7),01444\x1f\'9=82<.342\xff\xc0\x00\x0b\x08\x00\x01\x00\x01\x01\x01\x11\x00\xff\xc4\x00\x14\x00\x01\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x08\xff\xc4\x00\x14\x10\x01\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\xff\xda\x00\x08\x01\x01\x00\x00\x3f\x00\xd2\xff\xd9')
    print('Image created')
    pres = ppApp.Presentations.Add(0)
    print('Presentation created without window')
    slide = pres.Slides.Add(1, 12)
    print('Slide added')
    shp = slide.Shapes.AddPicture(FileName=img_path, LinkToFile=0, SaveWithDocument=-1, Left=0, Top=0, Width=100, Height=100)
    print('Picture added')
    pres.Close()
    print('Success with msoFalse!')
except Exception as e:
    print('Error:', e)
finally:
    if os.path.exists('test_img.jpg'): os.remove('test_img.jpg')
