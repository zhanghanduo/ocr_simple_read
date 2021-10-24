# from PIL.Image import Image
import os
from pdf2image import convert_from_path
import cv2
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'



# pdf_path = "./test/doc01003920210702154645.pdf"
# pdf_path = "./test/doc01004220210702154752.pdf"
# pdf_path = "./test/doc01324920211022183848.pdf"
pdf_path = "./test/doc01324720211022183816.pdf"

file_name = os.path.basename(pdf_path)
pages = convert_from_path(pdf_path=pdf_path, dpi=350)

image_name = "Page2_" + os.path.splitext(file_name)[0]+'.jpg'
pages[0].save(image_name, "JPEG")
im = cv2.imread(image_name)
# part1 = cv2.rectangle(im, (1245, 442), (1412, 488), color=(255, 0, 255), thickness=3)
# part2 = cv2.rectangle(im, (89, 546), (1060, 567), color=(255, 255, 0), thickness=3)
part1 = im[786:887, 2200:2455]
part2 = im[954:1022, 376:1872]
ret, thresh2 = cv2.threshold(part2, 120, 255, cv2.THRESH_BINARY)


part1_str = str(pytesseract.image_to_string(part1, config='--psm 7'))
part2_str = str(pytesseract.image_to_string(thresh2, config='--psm 7'))
print(part1_str)
print(part2_str)
os.remove(image_name)
