from PIL import Image
from pdf2image import convert_from_path
import pytesseract, re, os, glob, time, PyPDF2, csv, re

def crop(image, coordinates, savedName):
    croppedImage = image.crop(coordinates)
    croppedImage.save(savedName)

# Stored in (PO, Supplier, Date, Description, Invoice #, Amount) format
def process(docString, image, allPositions):
    # Filter looking for Vendor Identification Keywords
    if 'VERITIV OPERATING COMPANY' in docString:
        splitted = docString.split()
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'XPE01130'
        else:
            vendor = 'XPX001'
    elif 'INDIGO AMERICA INC' in docString:
        splitted = docString.split()
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'IND01130'
        else:
            vendor = 'IND001'
    else:
        print('Invoice not Recognized')
        return
    data = []
    # Get relevant position data from allPositions dictionary
    positions = allPositions[vendor]
    # PO
    crop(image, (positions[0][0], positions[0][1], positions[0][2], positions[0][3]), 'PO.png')
    PO = pytesseract.image_to_string(Image.open('PO.png'), config='--psm 7')
    if len(PO) == 5 and PO.isdigit():
        data.insert(0, PO)
    else:
        data.insert(0, '')
    # Supplier
    data.insert(1, vendor)
    # Date
    crop(image, (positions[1][0], positions[1][1], positions[1][2], positions[1][3]), 'date.png')
    date = pytesseract.image_to_string(Image.open('date.png'), config='--psm 7')
    data.insert(2, date)
    # Description
    crop(image, (positions[2][0], positions[2][1], positions[2][2], positions[2][3]), 'date.png')
    description = pytesseract.image_to_string(Image.open('date.png'), config='--psm 7')
    if(description != PO):
        data.insert(3, description)
    else:
        data.insert(3, '')
    # Invoice #
    crop(image, (positions[3][0], positions[3][1], positions[3][2], positions[3][3]), 'date.png')
    invoice = pytesseract.image_to_string(Image.open('date.png'), config='--psm 7')
    data.insert(4, invoice)
    # Amount
    crop(image, (positions[4][0], positions[4][1], positions[4][2], positions[4][3]), 'amount.png')
    amount = re.sub('[^\d.]', '', pytesseract.image_to_string(Image.open('amount.png'), config='--psm 7'))
    data.insert(5, amount)
    # Returned list
    return data

# Still need to add in functionality to edit the desired folder path
folder_path = 'C:/Users/Richard/Projects/OCR Invoice Processor/Invoices'
# Overall program start time
startTime = time.time()
# Storage for eventual CSV conversion
export = [['PO', 'Supplier', 'Date', 'Description', 'Invoice #', 'Amount']]
# Dictionary with all positional data (eventually implement import for this)
allPositions = {
    'XPX001':   [[467,846,746,902],[1186,265,1402,301],[467,846,746,902],[70,850,305,902],[1184,349,1403,386]],
    'XPE01130': [[467,846,746,902],[1186,265,1402,301],[467,846,746,902],[70,850,305,902],[1184,349,1403,386]],
    'IND001':   [[1178,323,1600,367],[1176,264,1355,300],[1177,391,1500,420],[1409,218,1518,247],[1176,1017,1375,1045]],
    'IND01130': [[1178,323,1600,367],[1176,264,1355,300],[1177,391,1500,420],[1409,218,1518,247],[1176,1017,1375,1045]],
    'ULN001':   [[200,712,315,740],[1406,205,1635,236],[],[],[]],
    'ULN01130': []
}
# Iterate through all files in folder_path
for filename in glob.glob(os.path.join(folder_path, '*.pdf')):
    # Iteration start time
    iterationStart = time.time()
    # Converting into image
    pages = convert_from_path(filename, single_file = True)
    for page in pages:
        page.save('temp.png', 'PNG')
    image = Image.open('temp.png')
    width, height = image.size
    crop(image, (0, 0, width, height / 2), 'cropped.png')
    # Read top half
    docString = pytesseract.image_to_string(Image.open('cropped.png'))
    # Check if Invoice
    if 'Invoice' in docString or 'INVOICE' in docString or 'invoice' in docString:
        export.append(process(docString, image, allPositions))
    else:
        print('not an invoice')
    print('Completed in ' + str(time.time() - iterationStart) + ' seconds')
csvfile = 'C:/Users/Richard/Projects/OCR Invoice Processor/output.csv'
# Only export as CSV if not empty
if len(export) > 6:
    with open(csvfile, "w", newline='') as output:
        writer = csv.writer(output)
        writer.writerows(export)
print('Total completed in ' + str(time.time() - startTime) + ' seconds')
