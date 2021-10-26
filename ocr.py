# from PIL.Image import Image
import os
from pdf2image import convert_from_path
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


pdf_path = "./test/doc01003920210702154645.pdf"
# pdf_path = "./test/doc01004220210702154752.pdf"
# pdf_path = "./test/doc01324920211022183848.pdf"
# pdf_path = "./test/doc01324720211022183816.pdf"

file_name = os.path.basename(pdf_path)
pages = convert_from_path(pdf_path=pdf_path, dpi=350)

image_name = "Page2_" + os.path.splitext(file_name)[0]+'.jpg'
im = pages[0]

area1 = (2172, 771, 2490, 859)
part1 = im.crop(area1)
area2 = (371, 943, 1800, 1081)
part2 = im.crop(area2)

part2.show()

options = "--psm 6 -c tessedit_char_blacklist=\/:*<>|?!"

part1_str = str(pytesseract.image_to_string(part1, config=options))
part2_str = str(pytesseract.image_to_string(part2, config=options)).partition('\n')[0]
print(part1_str)
print(part2_str)
# os.remove(image_name)
