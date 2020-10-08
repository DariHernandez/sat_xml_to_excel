#! python3
#Extract infomation from xml files and copy to the clipboard

import os, openpyxl, string, pprint
from xml.etree import ElementTree
from xml_info import getXmlEgresosInfo, getXmlIngresosInfo

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
    
    return currentRow

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

def writeTable (sheet, col, row, table): 
    """ Write all information of the table"""
    currentRow = row
    for seccion in table:
        currentRow = writeInfo (sheet, col=col, row=currentRow, data=seccion)



path = "/home/dari/Documentos/dari_developer_fact"
filePath = os.path.join (path, (os.path.basename (path) + ".xlsx"))
allInfo = []
ingresosFolder = "REGISTRO ANALÍTICO DE EGRESOS"
egresosFolder = "REGISTRO ANALÍTICO DE INGRESOS"

titlesIngreso = {'FECHA': [], 
                 'FOLIO': [],
                 'CLIENTE': [],
                 'IMPORTE': ['PÚBLICO G.', 'CLIENTES'],
                 'SUBTOTAL': [],
                 'IVA': [],
                 'TOTAL': []}

titlesEgreso = {'FECHA': [], 
                'IMPORTE': ['ING', 'EGR', 'PAG'],
                'FOLIO': [],
                'PROVEEDOR': [],
                'PAGADO': ['IMPORTE', 'IVA'],
                'DIFERIDO': ['IMPORTE', 'IVA'],
                'TOTAL': [],
                'COMENTARIOS': []}

formatedEgresoTitles = formatTitles (titlesEgreso)
formatedIngresoTitles = formatTitles (titlesIngreso)

# Open file o make new file and sheets
try: 
    wb = openpyxl.load_workbook(filePath)
except FileNotFoundError: 
    wb = openpyxl.Workbook()

for folder in os.listdir (path): 
    if os.path.isdir (os.path.join(path, folder)): 

        sheetName = folder

        # Open or make sheet
        if sheetName in wb.sheetnames: 
            # Replace or sheet with other name
            userContinue = input ('Sheet "%s" already exist.' % (sheetName) + \
            '¿Do you want to replace with the new xml info? (y/n)  ' )
            if userContinue.lower()[0] == "y":
                wb.remove (wb[sheetName])
                sheet = wb.create_sheet(sheetName)
            else: 
                counterSheet = 1
                newSheetName = sheetName + counterSheet
                while newSheetName in wb.sheetnames: 
                    counterSheet += 1
                sheet = wb.create_sheet(newSheetName)
        elif len(wb.sheetnames) == 1: 
            # Reaame the only sheet of the new file
            sheet = wb.active
            sheet.title = sheetName
        else: 
            sheet = wb.create_sheet(sheetName)
        
        # Seach and process files
        for subfolder in os.listdir (os.path.join (path, folder)):
            if os.path.isdir (os.path.join(path, folder, subfolder)): 
                table = []
                table.append([[subfolder]])
                if subfolder == "REGISTRO ANALÍTICO DE EGRESOS": 
                    # Write table
                    table.append (formatedEgresoTitles)
                    table.append (getXmlEgresosInfo (os.path.join(path, folder, subfolder)))
                    writeTable (sheet, col=11, row=1, table=table)
                    print ('XML files information written in "%s" sheet, "%s" table.' % (sheetName, subfolder))

                elif subfolder == "REGISTRO ANALÍTICO DE INGRESOS": 
                    # Write table
                    table.append (formatedIngresoTitles)
                    table.append (getXmlIngresosInfo (os.path.join(path, folder, subfolder)))
                    writeTable (sheet, col=1, row=1, table=table)
                    print ('XML files information written in "%s" sheet, "%s" table.' % (sheetName, subfolder))

print ("File '%s' saved." % (filePath))
wb.save (filePath)
