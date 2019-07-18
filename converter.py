import os, glob
from PIL import Image

from pdf2image import convert_from_path


folder_path = 'C:/Users/Richard/Projects/OCR Invoice Processor/temp'
count = 1

for filename in glob.glob(os.path.join(folder_path, '*.pdf')):
    pages = convert_from_path(filename)
    for page in pages:
        page.save('C:/Users/Richard/Projects/OCR Invoice Processor/temp/img' + str(count) + '.png', 'PNG')
        count += 1
    image = Image.open('temp.png')
