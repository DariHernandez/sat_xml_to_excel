#! python3
#Extract infomation from xml files and copy to the clipboard

import os, openpyxl, string, pprint
from xml.etree import ElementTree
from xml_info import getXmlEgresosInfo, getXmlIngresosInfo

def writeInfo (sheet, col, row, data):
    """ Write data in a wb sheet"""
    currentRow = row

    for dataLine in data: 
        currentCol = col 
        for dataItem in dataLine:
            sheet[chars[currentCol-1] + str(currentRow)] = dataItem
            currentCol += 1
        currentRow += 1 
    
    return currentRow

def writeTitles (sheet, col, row, dataTitles): 
    """ Write the titles in the ss"""
    currentCol = col
    currentRow = row

    for title, subtitles in dataTitles.items():
        # Write title
        titleCell = chars[currentCol-1] + str(currentRow)
        sheet[titleCell] = title
        
        if subtitles: 
            for subtitle in subtitles: 
                if subtitles.index(subtitle) > 0: 
                    currentCol += 1
                subtitleCell = chars[currentCol - 1] + str(currentRow + 1)
                sheet [subtitleCell] = subtitle
            endCell = chars[currentCol - 1] + str(currentRow)
            
            # Merge title cells
            sheet.merge_cells(str(titleCell + ":"+ endCell))
        else: 
            endCell = chars[currentCol - 1] + str(currentRow + 1)
            sheet.merge_cells(titleCell + ":"+  endCell)
        currentCol += 1

def writeTable (sheet, col, row, table): 
    """ Write all information of the table"""
    currentRow = row
    for seccion in table:
        currentRow = writeInfo (sheet, col=col, row=currentRow, data=seccion)

def writeTotals (sheet, col, row, numOfColumns, data): 
    """ Write the total values of each table"""
    currentRow = row
    for row in data: 
        currentRow += 1
    
    for addColumn in range (numOfColumns): 
        currentColumn = col + addColumn - 1 
        sumFormula = "=SUM(%s1:%s)" % (chars[currentColumn], chars[currentColumn] + str(currentRow-1))
        sheet [chars[currentColumn] + str(currentRow)] = sumFormula

def getSheet (wb, sheetName): 
    """ Select or make sheet """
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
    
    return sheet

def insertLogo (logo, anchor, width, height): 
    """ Paste logo in each page"""
    # Paste logo
    img = openpyxl.drawing.image.Image(logo)
    img.anchor = anchor
    img.width = width
    img.height = height
    sheet.add_image(img)

def writeMergeCells (sheet, colStart, rowStart, colNum, rowNum, text): 
    """ Write info in cell and merge"""
    textCell = chars[colStart-1] + str(rowStart)
    sheet[textCell] = text
    mergeCells = chars[colStart-1] + str(rowStart) + ":" + chars[colStart + colNum - 2] + str(rowStart + rowNum -1)
    sheet.merge_cells(mergeCells)
        

path = "/home/dari/Documentos/dari_developer_fact"
filePath = os.path.join (path, (os.path.basename (path) + ".xlsx"))
logoPath = os.path.join ("/home/dari/Projects/python/04 excel, xml and csv files/sat_xml_to_excel/logo.png")
allInfo = []
ingresosFolder = "REGISTRO ANALÍTICO DE EGRESOS"
egresosFolder = "REGISTRO ANALÍTICO DE INGRESOS"
columnIngresos = 1
columnEgresos = 10
chars = list(string.ascii_lowercase)

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

# Open file o make new file and sheets
try: 
    wb = openpyxl.load_workbook(filePath)
except FileNotFoundError: 
    wb = openpyxl.Workbook()

for folder in os.listdir (path): 
    if os.path.isdir (os.path.join(path, folder)): 

        sheetName = folder
        sheet = getSheet (wb, sheetName)

        # Seach and process files
        for subfolder in os.listdir (os.path.join (path, folder)):
            if os.path.isdir (os.path.join(path, folder, subfolder)): 
                # Write titles and info in merge cells
                #writeMergeCells (sheet, colStart, rowStart, colNum, rowNum, text)
                tableHeader = os.path.basename(path)
                writeMergeCells (sheet, columnIngresos, 1, 8, 1, tableHeader)
                writeMergeCells (sheet, columnIngresos, 2, 8, 1, subfolder)
                writeMergeCells (sheet, columnEgresos, 1, 11, 1, tableHeader)
                writeMergeCells (sheet, columnEgresos, 2, 11, 1, subfolder)
                if "egresos" in subfolder.lower(): 
                    # Titles
                    writeTitles (sheet, columnEgresos, 3, titlesEgreso)

                    # Write table
                    data = getXmlEgresosInfo (os.path.join(path, folder, subfolder))
                    writeInfo (sheet, col=columnEgresos, row=5, data=data)

                    # Add image
                    insertLogo (logoPath, 'J1', 50, 50)

                    # Totals
                    writeTotals (sheet, columnEgresos + 6, 5, 5, data)

                elif "ingresos" in subfolder.lower():  
                    # Titles
                    writeTitles (sheet, columnIngresos, 3, titlesIngreso)

                    # Write table
                    data = getXmlIngresosInfo (os.path.join(path, folder, subfolder))
                    writeInfo (sheet, col=columnIngresos, row=5, data=data)

                    # Add image
                    insertLogo (logoPath, 'A1', 50, 50)

                    # Totals
                    writeTotals (sheet, columnIngresos + 3, 5, 5, data)

                print ('XML files information written in "%s" sheet, "%s" table.' % (sheetName, subfolder))
print ("File '%s' saved." % (filePath))
wb.save (filePath)
