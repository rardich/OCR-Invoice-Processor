from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfFileReader
from positions import allPositions
import pytesseract, re, os, glob, time, csv, re

def crop(image, coordinates, savedName):
    croppedImage = image.crop(coordinates)
    croppedImage.save(savedName)

# Separate method necessary because amount listed on final page
def processLast(docString, image, allPositions, filename, vendor):
    # Process everything except amount on the first page
    data = []
    # Get relevant position data from allPositions dictionary
    positions = allPositions[vendor]
    # PO
    crop(image, (positions[0][0], positions[0][1], positions[0][2], positions[0][3]), 'PO.png')
    PO = re.sub('[^\d]', '', pytesseract.image_to_string(Image.open('PO.png'), config='--psm 7'))
    if len(PO) == 5 and PO.isdigit():
        data.insert(0, PO)
    else:
        PO = ''
        data.insert(0, '')
    # Supplier
    data.insert(1, vendor)
    # Date
    crop(image, (positions[1][0], positions[1][1], positions[1][2], positions[1][3]), 'date.png')
    date = re.sub('[^\d/]', '', pytesseract.image_to_string(Image.open('date.png'), config='--psm 7'))
    data.insert(2, date)
    # Description
    crop(image, (positions[2][0], positions[2][1], positions[2][2], positions[2][3]), 'desc.png')
    description = pytesseract.image_to_string(Image.open('desc.png'), config='--psm 7')
    if PO == '':
        data.insert(3, description)
    else:
        data.insert(3, '')
    # Invoice #
    crop(image, (positions[3][0], positions[3][1], positions[3][2], positions[3][3]), 'inv.png')
    invoice = pytesseract.image_to_string(Image.open('inv.png'), config='--psm 7')
    data.insert(4, invoice)
    # Make new image for last page with amount on it
    pages = convert_from_path(filename)
    for page in pages:
        page.save('temp.png', 'PNG')
    image = Image.open('temp.png')
    # Amount
    crop(image, (positions[4][0], positions[4][1], positions[4][2], positions[4][3]), 'amount.png')
    amount = re.sub('[^\d.]', '', pytesseract.image_to_string(Image.open('amount.png'), config='--psm 7'))
    data.insert(5, amount)
    # Return list
    print(data)
    return data


def process(docString, image, allPositions, filename):
    print(filename)
    splitted = docString.split()
    # Filter looking for Vendor Identification Keywords
    if 'INDIGO AMERICA INC' in docString:
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'IND01130'
        else:
            vendor = 'IND001'
    elif 'IA' in splitted and 'THOMPSON' in splitted:
        # LBS amount has dynamic location, so will look for it in text instead
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'LBS01130'
        else:
            vendor = 'LBS010'
    elif 'Midland Paper Company' in docString:
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'MID01130'
        else:
            vendor = 'MDL001'
        pdf = PdfFileReader(filename, 'rb')
        numPages = pdf.getNumPages() 
        if numPages > 1:
            return processLast(docString, image, allPositions, filename, vendor)
    elif 'Norcross' in docString and 'Corporate' in docString:
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        vendor = 'SPO001'
    elif 'ULINE' in docString:
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'ULN01130'
        else:
            vendor = 'ULN001'
        pdf = PdfFileReader(filename, 'rb')
        numPages = pdf.getNumPages() 
        if numPages > 1:
            return processLast(docString, image, allPositions, filename, vendor)
    elif 'VERITIV OPERATING COMPANY' in docString:
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'XPE01130'
        else:
            vendor = 'XPX001'
    else:
        print('Invoice not Recognized')
        print(splitted)
        return
    data = []
    # Get relevant position data from allPositions dictionary
    positions = allPositions[vendor]
    # PO
    crop(image, (positions[0][0], positions[0][1], positions[0][2], positions[0][3]), 'PO.png')
    PO = re.sub('[^\d]', '', pytesseract.image_to_string(Image.open('PO.png'), config='--psm 7'))
    if len(PO) == 5 and PO.isdigit():
        data.insert(0, PO)
    else:
        PO = ''
        data.insert(0, '')
    # Supplier
    data.insert(1, vendor)
    # Date
    crop(image, (positions[1][0], positions[1][1], positions[1][2], positions[1][3]), 'date.png')
    date = re.sub('[^\d/]', '', pytesseract.image_to_string(Image.open('date.png'), config='--psm 7'))
    data.insert(2, date)
    # Description
    crop(image, (positions[2][0], positions[2][1], positions[2][2], positions[2][3]), 'desc.png')
    description = pytesseract.image_to_string(Image.open('desc.png'), config='--psm 7')
    if PO == '':
        data.insert(3, description)
    else:
        data.insert(3, '')
    # Invoice #
    crop(image, (positions[3][0], positions[3][1], positions[3][2], positions[3][3]), 'inv.png')
    invoice = pytesseract.image_to_string(Image.open('inv.png'), config='--psm 7')
    data.insert(4, invoice)
    # Amount
    crop(image, (positions[4][0], positions[4][1], positions[4][2], positions[4][3]), 'amount.png')
    amount = re.sub('[^\d.]', '', pytesseract.image_to_string(Image.open('amount.png'), config='--psm 7'))
    data.insert(5, amount)
    # Invoice # Exceptions
    if vendor == 'SPO001':
        tempInv = list(data[4])
        tempInv[0] = 'I'
        data[4]= ''.join(tempInv)
    # Amount Exceptions
    if vendor == 'LBS01130' or vendor == 'LBS010':
        amount = re.sub('[^\d.]', '', splitted[-4])
        data[5] = amount
    elif vendor == 'SPO001':
        amount = re.sub('[^\d.]', '', splitted[-1])
        data[5] = amount
    # Return list
    print(data)
    return data

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
        pdf = open(filename, 'rb')
    # Check if invoice
    if 'Invoice' in docString or 'INVOICE' in docString or 'invoice' in docString:
        data = process(docString, image, allPositions, filename)
        if data is not None:
            export.append(data)
    else:
        print('not an invoice')
    count += 1
    print('Completed in ' + str(time.time() - iterationStart) + ' seconds')
# Only export as CSV if not empty
csvfile = 'C:/Users/Richard/Projects/OCR Invoice Processor/output.csv'
# If nothing besides the labels, len() will read the length of the list as 6, if there is content
# the length will be of the number of lists inside the list instead
if len(export) > 0 and len(export) != 6:
    with open(csvfile, "w+", newline='') as output:
        writer = csv.writer(output)
        writer.writerows(export)
print(str(count) + ' invoices completed in ' + str(time.time() - startTime) + ' seconds')
print('Averaged ' + str((time.time() - startTime)/count) + ' seconds')
