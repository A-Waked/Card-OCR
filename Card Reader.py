# Python bytecode 3.6 (3372)
# Embedded file name: main.py
# Decompiled by https://python-decompiler.com
nltk_packages = [
 'stopwords', 'punkt', 'averaged_perceptron_tagger', 'maxent_ne_chunker', 'words']
print('Ensuring all packages are up to date...')
import image_slicer, re, os, nltk
for package in nltk_packages:
    nltk.download(package)

import xlsxwriter, win32com.client as win32
from nltk.chunk.named_entity import *
from nltk.corpus import stopwords
from random import *
try:
    import Image
except ImportError:
    from PIL import Image

import pytesseract
stop = stopwords.words('english')
pytesseract.pytesseract.tesseract_cmd = 'OCR-ENGINE/tesseract'
print('\n-----------------------------------------------------------------------------------------------------------------------------------------------------------------------\nCopyright 2017 ABDUL WAKED\n\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the \nSoftware without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, \nand to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n\nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n\nTHE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF \nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR \nANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH\nTHE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.\n-----------------------------------------------------------------------------------------------------------------------------------------------------------------------\n')

def myround(x, base):
    return int(base * round(float(x) / base))


def myround(x, base):
    return int(base * round(float(x) / base))


def extract_phone_numbers(string):
    r = re.compile('(\\d{3}[-\\.\\s]??\\d{3}[-\\.\\s]??\\d{4}|\\(\\d{3}\\)\\s*\\d{3}[-\\.\\s]??\\d{4}|\\d{3}[-\\.\\s]??\\d{4})')
    phone_numbers = r.findall(string)
    return [re.sub('\\D', '', number) for number in phone_numbers]


def extract_email_addresses(string):
    r = re.compile('[\\w\\.-]+@[\\w\\.-]+')
    return r.findall(string)


def extract_web_addresses(string):
    r = re.compile('(?i)\\b((?:https?:(?:/{1,3}|[a-z0-9%])|[a-z0-9.\\-]+[.](?:com|net|org|edu|gov|mil|aero|asia|biz|cat|coop|info|int|jobs|mobi|museum|name|post|pro|tel|travel|xxx|ac|ad|ae|af|ag|ai|al|am|an|ao|aq|ar|as|at|au|aw|ax|az|ba|bb|bd|be|bf|bg|bh|bi|bj|bm|bn|bo|br|bs|bt|bv|bw|by|bz|ca|cc|cd|cf|cg|ch|ci|ck|cl|cm|cn|co|cr|cs|cu|cv|cx|cy|cz|dd|de|dj|dk|dm|do|dz|ec|ee|eg|eh|er|es|et|eu|fi|fj|fk|fm|fo|fr|ga|gb|gd|ge|gf|gg|gh|gi|gl|gm|gn|gp|gq|gr|gs|gt|gu|gw|gy|hk|hm|hn|hr|ht|hu|id|ie|il|im|in|io|iq|ir|is|it|je|jm|jo|jp|ke|kg|kh|ki|km|kn|kp|kr|kw|ky|kz|la|lb|lc|li|lk|lr|ls|lt|lu|lv|ly|ma|mc|md|me|mg|mh|mk|ml|mm|mn|mo|mp|mq|mr|ms|mt|mu|mv|mw|mx|my|mz|na|nc|ne|nf|ng|ni|nl|no|np|nr|nu|nz|om|pa|pe|pf|pg|ph|pk|pl|pm|pn|pr|ps|pt|pw|py|qa|re|ro|rs|ru|rw|sa|sb|sc|sd|se|sg|sh|si|sj|Ja|sk|sl|sm|sn|so|sr|ss|st|su|sv|sx|sy|sz|tc|td|tf|tg|th|tj|tk|tl|tm|tn|to|tp|tr|tt|tv|tw|tz|ua|ug|uk|us|uy|uz|va|vc|ve|vg|vi|vn|vu|wf|ws|ye|yt|yu|za|zm|zw)/)(?:[^\\s()<>{}\\[\\]]+|\\([^\\s()]*?\\([^\\s()]+\\)[^\\s()]*?\\)|\\([^\\s]+?\\))+(?:\\([^\\s()]*?\\([^\\s()]+\\)[^\\s()]*?\\)|\\([^\\s]+?\\)|[^\\s`!()\\[\\]{};:\'".,<>?])|(?:(?<!@)[a-z0-9]+(?:[.\\-][a-z0-9]+)*[.](?:com|net|org|edu|gov|mil|aero|asia|biz|cat|coop|info|int|jobs|mobi|museum|name|post|pro|tel|travel|xxx|ac|ad|ae|af|ag|ai|al|am|an|ao|aq|ar|as|at|au|aw|ax|az|ba|bb|bd|be|bf|bg|bh|bi|bj|bm|bn|bo|br|bs|bt|bv|bw|by|bz|ca|cc|cd|cf|cg|ch|ci|ck|cl|cm|cn|co|cr|cs|cu|cv|cx|cy|cz|dd|de|dj|dk|dm|do|dz|ec|ee|eg|eh|er|es|et|eu|fi|fj|fk|fm|fo|fr|ga|gb|gd|ge|gf|gg|gh|gi|gl|gm|gn|gp|gq|gr|gs|gt|gu|gw|gy|hk|hm|hn|hr|ht|hu|id|ie|il|im|in|io|iq|ir|is|it|je|jm|jo|jp|ke|kg|kh|ki|km|kn|kp|kr|kw|ky|kz|la|lb|lc|li|lk|lr|ls|lt|lu|lv|ly|ma|mc|md|me|mg|mh|mk|ml|mm|mn|mo|mp|mq|mr|ms|mt|mu|mv|mw|mx|my|mz|na|nc|ne|nf|ng|ni|nl|no|np|nr|nu|nz|om|pa|pe|pf|pg|ph|pk|pl|pm|pn|pr|ps|pt|pw|py|qa|re|ro|rs|ru|rw|sa|sb|sc|sd|se|sg|sh|si|sj|Ja|sk|sl|sm|sn|so|sr|ss|st|su|sv|sx|sy|sz|tc|td|tf|tg|th|tj|tk|tl|tm|tn|to|tp|tr|tt|tv|tw|tz|ua|ug|uk|us|uy|uz|va|vc|ve|vg|vi|vn|vu|wf|ws|ye|yt|yu|za|zm|zw)\\b/?(?!@)))')
    return r.findall(string)


def ie_preprocess(document):
    document = (' ').join([i for i in document.split() if i not in stop])
    sentences = nltk.sent_tokenize(document)
    sentences = [nltk.word_tokenize(sent) for sent in sentences]
    sentences = [nltk.pos_tag(sent) for sent in sentences]
    return sentences


def extract_names(document):
    names = []
    sentences = ie_preprocess(document)
    for tagged_sentence in sentences:
        for chunk in nltk.ne_chunk(tagged_sentence):
            if type(chunk) == nltk.tree.Tree:
                if chunk.label() == 'PERSON':
                    names.append((' ').join([c[0] for c in chunk]))

    return names


def OCR(image):
    print('Starting OCR Process')
    
    im = image
    
    image_x = im.size[0]
    image_y = im.size[1]
    
    new_image_x = myround(im.size[0], 594)
    new_image_y = myround(im.size[1], 1037)
    
    slices = new_image_x // 594 * (new_image_y // 1037)
    
    leftover_x = image_x - new_image_x
    leftover_y = image_y - new_image_y
    
    im = im.crop((0, 0, image_x - leftover_x, image_y - leftover_y)).save('card.png')
    
    tiles = image_slicer.slice('card.png', slices, False)
    
    excel_info = []
    
    for tile in tiles:
        card_ID = randint(100000, 999999)
        
        name = ('card {0}.png').format(card_ID)

        print('Processing:', name)
        
        tile.save(name)
        
        string = pytesseract.image_to_string(Image.open(name))
        numbers = (',').join(extract_phone_numbers(string))
        emails = (',').join(extract_email_addresses(string))
        names = (',').join(extract_names(string))
        sites = (',').join(extract_web_addresses(string))
        
        info = [
         card_ID, names, sites, emails, numbers, string]
        
        excel_info.append(info)

    print('Finished reading and seperating cards. Saving information...')
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    workbook = xlsxwriter.Workbook('cardInfo.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    excel_info = tuple(excel_info)
    row = 1
    col = 0
    print('Creating Excel file')
    worksheet.write('A1', 'Card ID', bold)
    worksheet.write('B1', 'Client Name', bold)
    worksheet.write('C1', 'Websites', bold)
    worksheet.write('D1', 'Emails', bold)
    worksheet.write('E1', 'Phone Numbers', bold)
    worksheet.write('F1', 'Unorganized Info', bold)
    print('Saving Information onto Excel file')
    for card_number, name, site, email, number, text in excel_info:
        worksheet.write(row, col, card_number)
        worksheet.write(row, col + 1, name)
        worksheet.write(row, col + 2, site)
        worksheet.write(row, col + 3, email)
        worksheet.write(row, col + 4, number)
        worksheet.write(row, col + 5, text)
        row += 1

    workbook.close()
    print('Formatting Excel File')
    dir_path = os.path.dirname(os.path.realpath(__file__))
    dir_path = dir_path.replace('\\', '/')
    dir_path = dir_path + '/cardInfo.xlsx'
    wb = excel.Workbooks.Open(dir_path)
    ws = wb.Worksheets('Sheet1')
    ws.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()
    print('Finished OCR process. Please be aware that cards may not scan properly if the card is not simply designed. Double check information on the excel file!')
    print('Or you can simply reference a card by using the generated pictures. The ID on the Excel file uses the same ID as the image file.')


OCR(Image.open('card.png'))
