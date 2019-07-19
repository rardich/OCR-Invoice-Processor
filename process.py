from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfFileReader
from positions import allPositions, months
import pytesseract, re, os, glob, time, csv, re, itertools

# Crops image based on coordinates and returns image
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
    if vendor != 'XPX001' or vendor != 'XPE01130':
        invoice = invoice.lstrip('0').strip('.')
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

    # Amount Exceptions
    if vendor == 'SIT001' or vendor == 'STR01130':
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        amount = re.sub('[^\d.]', '', splitted[-1])
        data[5] = amount
    elif vendor == 'UEI001':
        pdf = PdfFileReader(filename, 'rb')
        numPages = pdf.getNumPages() 
        page = pdf.getPage(numPages - 1)
        docString = page.extractText()
        splitted = docString.split()
        amount = re.sub('[^\d.]', '', splitted[splitted.index('Total:') + 2])
        data[5] = amount
    elif vendor == 'ZEE001':
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        amount = re.sub('[^\d.]', '', splitted[splitted.index('Grand') + 2])
        data[5] = amount

    # Return list
    print(data)
    return data

# Decides on vendor and fills out data
def process(docString, image, allPositions, filename):
    splitted = docString.split()
    data = []
    # Apple Courier
    if 'Apple Courier' in docString:
        vendor = 'APP01130'
    # BNBindery
    elif 'BRIDGEPORT NATIONAL' in docString:
        vendor = 'BNB001'
    # Cedar Grove
    elif 'CEDAR GROVE' in docString:
        vendor = 'CGO001'
    # Clearbags
    elif 'www.clearbags.com' in docString:
        vendor = 'CLE001'
        positions = allPositions[vendor]
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        PO = ''
        for word in splitted:
            if word[0] == '2' and len(word) == 5:
                PO = word
                break
        crop(image, (positions[1][0], positions[1][1], positions[1][2], positions[1][3]), 'date.png')
        date = re.sub('[^\d/]', '', pytesseract.image_to_string(Image.open('date.png'), config='--psm 7'))
        description = ''
        crop(image, (positions[3][0], positions[3][1], positions[3][2], positions[3][3]), 'inv.png')
        invoice = pytesseract.image_to_string(Image.open('inv.png'), config='--psm 7')
        amount = re.sub('[^\d.]', '', splitted[splitted.index('Due') + 1])
        if amount == '':
            crop(image, (positions[4][0], positions[4][1], positions[4][2], positions[4][3]), 'amount.png')
            amount = re.sub('[^\d.]', '', pytesseract.image_to_string(Image.open('amount.png'), config='--psm 7'))
        data.insert(0, PO)
        data.insert(1, vendor)
        data.insert(2, date)
        data.insert(3, description)
        data.insert(4, invoice)
        data.insert(5, amount)
        print(data)
        return data 
    # Crown Lift Trucks
    elif '!' in docString and '@' in docString and '&' in docString:
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        PO = ''
        for word in splitted:
            if word[0] == '2' and len(word) == 5:
                PO = word
                break
        vendor = 'CLT130'
        date = splitted[splitted.index('Date:') + 1]
        try:
            description = 'S/N:' + splitted[splitted.index('S/N:') + 1]
        except:
            description = 'Rental'
        invoice = splitted[splitted.index('Invoice:') + 1]
        amount = re.sub('[^\d.]', '', splitted[splitted.index('Due:') + 1])
        positions = allPositions[vendor]
        if amount == '':
            crop(image, (positions[4][0], positions[4][1], positions[4][2], positions[4][3]), 'amount.png')
            amount = re.sub('[^\d.]', '', pytesseract.image_to_string(Image.open('amount.png'), config='--psm 7'))
        data.insert(0, PO)
        data.insert(1, vendor)
        data.insert(2, date)
        data.insert(3, description)
        data.insert(4, invoice)
        data.insert(5, amount)
        print(data)
        return data    
    # Crystal Springs
    # Dwyer R+D
    elif 'Dwyer' in docString:
        vendor = 'DRD001'
    # Evergreen Vending
    # KNOWN BUG: NOT RECOGNIZING DATES ON OCCASION
    elif 'DATEINVOICE' in docString and 'NUMBERTERMSRoute' in docString:
        vendor = 'EVE001'
    # First Choice
    # Fujifilm
    # GP2
    elif 'www.gp2tech.com' in docString:
        if 'Big' in splitted and 'Shanty' in splitted:
            vendor = 'GP201130'
        else:
            vendor = 'GP2001'
    # GPA
    elif 'www.askgpa.com' in docString:
        if 'Big' in splitted and 'Shanty' in splitted:
            vendor = 'GPA01130'
        else:
            vendor = 'GPA001'
    # Grimco
    elif 'www.grimco.com' in docString:
        vendor = 'GRI01130'
    # HP Indigo
    elif 'INDIGO AMERICA INC' in docString:
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'IND01130'
        else:
            vendor = 'IND001'
    # Labels Direct
    # LBS
    elif 'IA' in splitted and 'THOMPSON' in splitted:
        # LBS amount has dynamic location, so will look for it in text instead
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'LBS01130'
        else:
            vendor = 'LBS010'
    # Livingston
    elif 'ITASCA' in docString and 'Pierce' in docString:
        vendor = 'LII001'
    # Mac Paper
    # MaxVision
    elif 'Kennestone' in docString and 'Circle' in docString:
        vendor = 'MAX01130'
    # Mckinney
    # Metro Trailer
    # Midland
    elif 'Midland Paper Company' in docString:
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'MID01130'
        else:
            vendor = 'MDL001'
        pdf = PdfFileReader(filename, 'rb')
        numPages = pdf.getNumPages() 
        if numPages > 1:
            return processLast(docString, image, allPositions, filename, vendor)
    # National Print
    elif 'National Print Wholesale' in docString:
        vendor = 'NPW130'
    # Nobelus
    # Oracle
    elif 'Oracle America' in docString:
        vendor = 'ORA001'
    # Paper Handling
    elif 'Paper Handling Solutions' in docString:
        # Needs to be processed completely as a string because invoices are scanned into PDF
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        vendor = 'PHS01130'
        try:
            PO = splitted[splitted.index('ReiPreQl') + 1]
        except:
            try:
                PO = splitted[splitted.index('ReiPre0dl') + 1]
            except:
                try:
                    PO = splitted[splitted.index('ReiPredl') + 1]
                except:
                    try:
                        PO = splitted[splitted.index('ReiPreol') + 1]
                    except:
                        PO = splitted[splitted.index('ReiPre0l') + 1]
        if PO.isdigit() != True:
            description = PO + ' ' + splitted[splitted.index(PO) + 1]
            PO = ''
        else:
            description = ''
        try:
            date = re.sub('[^\d/]','', str(months[splitted[splitted.index('Date:') + 1]]) + '/' + splitted[splitted.index('Date:') + 2] + '/' + splitted[splitted.index('Date:') + 3])
        except:
            date = re.sub('[^\d/]','', str(months[splitted[splitted.index('Date.') + 1]]) + '/' + splitted[splitted.index('Date.') + 2] + '/' + splitted[splitted.index('Date.') + 3])
        try:
            invoice = splitted[splitted.index('Number:') + 1]
        except:
            invoice = splitted[splitted.index('Number.') + 1]
        amount = splitted[splitted.index('TOTAL') + 1]
        data.insert(0, PO)
        data.insert(1, vendor)
        data.insert(2, date)
        data.insert(3, description)
        data.insert(4, invoice)
        data.insert(5, amount)
        print(data)
        return data
    # Perimeter Office
    elif 'PERIMETER OFFICE PRODUCTS' in docString:
        vendor = 'POP01130'
    # Purolator
    elif 'www.purolatorinternational.com' in docString:
        vendor = 'PRL001'
    # S-One
    # Scott Lithographing
    # SHI International
    elif 'SHI International Corp' in docString:
        vendor = 'SHI001'
        pdf = PdfFileReader(filename, 'rb')
        numPages = pdf.getNumPages() 
        if numPages > 1:
            return processLast(docString, image, allPositions, filename, vendor)
    # Sound Maintenance
    # Spoke
    elif 'Norcross' in docString and 'Corporate' in docString:
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        vendor = 'SPO001'
    # Structured
    elif 'CLACKAMAS' in docString and '400' in docString:
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'STR01130'
        else:
            vendor = 'SIT001'
        pdf = PdfFileReader(filename, 'rb')
        numPages = pdf.getNumPages() 
        if numPages > 1:
            return processLast(docString, image, allPositions, filename, vendor)
    # TEC Lighting
    # Tradebe
    elif 'Tradebe' in docString:
        docString = pytesseract.image_to_string(image)
        splitted = docString.split()
        vendor = 'TRA01130'
    # Uline
    elif 'ULINE' in docString:
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'ULN01130'
        else:
            vendor = 'ULN001'
        pdf = PdfFileReader(filename, 'rb')
        numPages = pdf.getNumPages() 
        if numPages > 1:
            return processLast(docString, image, allPositions, filename, vendor)
    # Universal Engraving
    elif 'www.ueigroup.com' in docString:
        vendor = 'UEI001'
        pdf = PdfFileReader(filename, 'rb')
        numPages = pdf.getNumPages() 
        if numPages > 1:
            return processLast(docString, image, allPositions, filename, vendor)
    # USADATA
    # Veritiv
    elif 'VERITIV OPERATING COMPANY' in docString:
        if 'BIG' in splitted and 'SHANTY' in splitted:
            vendor = 'XPE01130'
        else:
            vendor = 'XPX001'
    # Washington Alarm
    elif 'Washington Alarm' in docString:
        vendor = 'WAI001'
    # Waste Not Paper
    elif 'Please RUSH this freight collect order!!' in docString:
        vendor = 'WNP001'
    # Zee Medical
    elif 'Zee Medical' in docString:
        vendor = 'ZEE001'
        pdf = PdfFileReader(filename, 'rb')
        numPages = pdf.getNumPages() 
        if numPages > 1:
            return processLast(docString, image, allPositions, filename, vendor)
    # Other
    else:
        print('Invoice not Recognized')
        print(splitted)
        return
    # Get relevant position data from allPositions dictionary
    positions = allPositions[vendor]
    # PO
    crop(image, (positions[0][0], positions[0][1], positions[0][2], positions[0][3]), 'PO.png')
    PO = re.sub('[^\d]', '', pytesseract.image_to_string(Image.open('PO.png'), config='--psm 7'))
    if len(PO) == 5 and PO[0] == '2':
        data.insert(0, PO)
    else:
        PO = ''
        data.insert(0, '')
    # Supplier
    data.insert(1, vendor)
    # Date
    crop(image, (positions[1][0], positions[1][1], positions[1][2], positions[1][3]), 'date.png')
    date = re.sub('[^\d/]', '', pytesseract.image_to_string(Image.open('date.png'), config='--psm 7'))
    if vendor == 'LII001':
        tempDate = date[3:6] + date[0:3] + date[6:]
        date = tempDate
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
    if vendor != 'XPX001' or vendor != 'XPE01130':
        invoice = invoice.lstrip('0').strip('.')
    data.insert(4, invoice)
    # Amount
    crop(image, (positions[4][0], positions[4][1], positions[4][2], positions[4][3]), 'amount.png')
    amount = re.sub('[^\d.]', '', pytesseract.image_to_string(Image.open('amount.png'), config='--psm 7'))
    data.insert(5, amount)
    
    # Date Exceptions
    if vendor == 'PRL001':
        crop(image, (positions[1][0], positions[1][1], positions[1][2], positions[1][3]), 'date.png')
        date = pytesseract.image_to_string(Image.open('date.png'), config='--psm 7')
        date = re.sub('[^\d/]','', str(months[date[0:3]]) + '/' + str(date[-4]) + '/' + str(date[-8:-6]))
        data[2] = date
    elif vendor == 'GRI01130':
        crop(image, (positions[1][0], positions[1][1], positions[1][2], positions[1][3]), 'date.png')
        date = pytesseract.image_to_string(Image.open('date.png'), config='--psm 7')
        date = re.sub('[^\d/]','', str(months[date[3:6]]) + '/' + str(date[0:2]) + '/' + str(date[-4:]))
        data[2] = date
    # Description Exceptions
    if vendor == 'APP01130' or vendor == 'CGO001':
        description = 'Svc ' + description
        data[3] = description
    elif vendor == 'NPW130':
        description = ''
        data[3] = description
    elif vendor == 'EVE001':
        if description != 'Avanti Subsidy':
            description = 'Coffee'
        data[3] = description
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
    elif vendor == 'TRA01130':
        amount = splitted[-12]
        if len(amount) == 10 and amount.isdigit() != True:
            amount = re.sub('[^\d.]', '', splitted[-13])
        else:
            amount = re.sub('[^\d.]', '', amount)
        data[5] = amount
    elif vendor == 'UEI001':
        amount = re.sub('[^\d.]', '', splitted[splitted.index('Total:') + 2])
        data[5] = amount
    elif vendor == 'DRD001':
        amount = re.sub('[^\d.]', '', splitted[-1])
        data[5] = amount
    # Return list
    print(data)
    return data