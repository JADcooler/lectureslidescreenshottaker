from PIL import ImageGrab
import keyboard
import time
from docx import Document
from datetime import date
from docx.shared import Inches



document = Document()
today=date.today()
while True:
    
    if keyboard.is_pressed('z'):
        image = ImageGrab.grab((0,100,1280,720))
        image.save("1.png")
        time.sleep(1)
        document.add_picture('1.png',width=Inches(7.0))
        document.save(today.strftime('%d-%m-%y.docx'))
    else:
        pass
    
