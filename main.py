from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfFileReader
from positions import allPositions, months
from process import crop, process
import pytesseract, re, os, glob, time, csv, re, itertools

# Still need to add in functionality to edit the desired folder path
folder_path = 'C:/Users/Richard/Projects/OCR Invoice Processor/Invoices'
# Overall program start time
startTime = time.time()
# Storage for eventual CSV conversion
export = [['PO', 'Supplier', 'Date', 'Description', 'Invoice #', 'Amount']]
count = 0

# Iterate through all files in folder_path
for filename in glob.glob(os.path.join(folder_path, '*.pdf')):
    # Iteration start time
    iterationStart = time.time()
    print(filename)
    # Converting into image
    pages = convert_from_path(filename, single_file = True)
    for page in pages:
        page.save('temp.png', 'PNG')
    image = Image.open('temp.png')
    # Extracting the text to check for vendor and if invoice
    # Try and except used because of exception: ValueError: 
    # invalid literal for int() with base 10: 'obj' not allowing 
    # certain PDF file names to be used
    try:
        pdf = PdfFileReader(filename, 'rb')
        page = pdf.getPage(0)
        docString = page.extractText()
    except:
        image = Image.open('temp.png')
        width, height = image.size
        crop(image, (0, 0, width, height / 2), 'cropped.png')
        # Read top half
        docString = pytesseract.image_to_string(Image.open('cropped.png'))
    if docString == '' or docString.isspace():
        image = Image.open('temp.png')
        width, height = image.size
        crop(image, (0, 0, width, height / 2), 'cropped.png')
        # Read top half
        docString = pytesseract.image_to_string(Image.open('cropped.png'))
    # Check if invoice
    if 'Invoice' in docString or 'INVOICE' in docString or 'invoice' in docString or ('!' in docString and '@' in docString and '&' in docString):
        data = process(docString, image, allPositions, filename)
        if data is not None:
            export.append(data)
    else:
        print(docString)
        print('not an invoice')
    count += 1
    print('Completed in ' + str(time.time() - iterationStart) + ' seconds')
# Only export as CSV if not empty
csvfile = 'C:/Users/Richard/Projects/OCR Invoice Processor/output.csv'
# If nothing besides the labels, len() will read the length of the list as 6, if there is content
# the length will be of the number of lists inside the list instead
if len(export) > 0 and len(list(itertools.chain.from_iterable(export))) != 6:
    with open(csvfile, "w+", newline='') as output:
        writer = csv.writer(output)
        writer.writerows(export)
print(str(count) + ' invoices completed in ' + str(time.time() - startTime)  + ' seconds')
if count != 0:
    print('Averaged ' + str((time.time() - startTime)/count) + ' seconds')
