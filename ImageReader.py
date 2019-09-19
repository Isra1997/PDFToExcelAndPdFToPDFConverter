# reading image and converting to text alamiir.ashraf@gmail.com
from pytesseract import image_to_string
from PIL import Image
import PyPDF2

image=Image.open("/Users/israragheb/Desktop/out.jpg")
file=open('PDFtest.txt','a')
file.write(image_to_string(image))
file.close()
