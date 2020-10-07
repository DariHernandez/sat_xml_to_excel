#! python3
#Extract infomation from xml files and copy to the clipboard

import os, openpyxl, string, pprint
from xml.etree import ElementTree
from xml_info import getXmlEgresosInfo

def writeInfo (sheet, col, row, data):
    """ Write data in a wb sheet"""
    currentRow = row
    chars = list(string.ascii_lowercase)

    for dataLine in data: 
        currentCol = col 
        for dataItem in dataLine:
            sheet[chars[currentCol-1] + str(currentRow)] = dataItem
            currentCol += 1
        currentRow += 1 

def formatTitles (currentTitles): 
    """ Format dicc of titles for the spreadsheet table"""
    titlesFormated = []
    titles = []
    subtitles = []

    for key, values in currentTitles.items():
        titles.append (key)
        if values: 
            for value in values: 
                if values.index (value) > 0: 
                    titles.append(' ')
                subtitles.append (value)
        else: 
            subtitles.append (' ')
    titlesFormated.append (titles)
    titlesFormated.append (subtitles)
    return titlesFormated

path = "/home/dari/Documentos/dari_developer_fact"
allInfo = []

titlesEgreso = {'FECHA': [], 
                'EFECTO': ['ING', 'EGR', 'PAG'],
                'FOLIO': [],
                'PROVEEDOR': [],
                'PAGADO': ['IMPORTE', 'IVA'],
                'DIFERIDO': ['IMPORTE', 'IVA'],
                'TOTAL': [],
                'COMENTARIOS': []}

formatedEgresoTitles = formatTitles (titlesEgreso)

for folder in os.listdir (path):
    if os.path.isdir (os.path.join(path, folder)):  
        filePath = os.path.join (path, (os.path.basename (path) + ".xlsx"))

        # Open file o make new file and sheets
        sheetName = folder
        try: 
            wb = openpyxl.load_workbook(filePath)
            # Use or make sheet
            if not sheetName in wb.sheetnames: 
                sheet = wb.create_sheet(folder)
            else: 
                sheet = wb[sheetName]
        except FileNotFoundError: 
            wb = openpyxl.Workbook()
            # Make sheet
            sheet = wb.active
            sheet.title = sheetName

        if not sheetName in wb.sheetnames: 
            sheet = wb.create_sheet(folder)
        else: 
            sheet = wb[sheetName]
        
        formatedShortedInfo = getXmlEgresosInfo (os.path.join(path, folder))

        writeInfo (sheet, 1, 1, formatedEgresoTitles)
        writeInfo (sheet, 1, 3, formatedShortedInfo)
        print ("XMl info is now in '%s' sheet." % (sheetName))

        wb.save (filePath)
        print ("File '%s' saved." % (filePath))
        
