#! python3
# Format the xlsx file

import openpyxl, string
from openpyxl.styles import Font, Alignment, PatternFill

chars = list(string.ascii_lowercase)

def setWidthColumn (sheet, currentCell, currentCol): 
    """ Set the with for each column """

    if currentCell.value == "X":
        width = len(str(currentCell.value)) + 5
    else: 
        width = len(str(currentCell.value)) + 2

    if currentCell.value == "X" or width > sheet.column_dimensions[chars[currentCol-1]].width: 
        sheet.column_dimensions[chars[currentCol-1]].width = width


def formatData (filePath, sheetName, colStart, rowStart, colNum, rowNum): 

    wb = openpyxl.load_workbook(filePath) 
    sheet = wb [sheetName]

    table = chars[colStart-1].upper() + str(rowStart) + ":" + chars[colStart + colNum -1].upper() + str(rowStart + rowNum)
    print ('Applying styles to table %s' % (table))

    # Get and format cells
    for currentRow in range (rowStart, rowStart + rowNum): 
        for currentCol in range (colStart, colStart + colNum):
            titleCell = "%s%s" % (chars[currentCol-1], rowStart - 2)
            subtitleCell = "%s%s" % (chars[currentCol-1], rowStart - 1)
            currentCellName = "%s%s" % (chars[currentCol-1], currentRow)
            currentCell = sheet[currentCellName]

            titleName = str(sheet[titleCell].value).lower().strip()
            subtitleName = str(sheet[subtitleCell].value).lower().strip()

            # Set width for all +columns
            setWidthColumn (sheet, currentCell, currentCol)

            print (titleName, subtitleName, currentCellName)

            if titleName == "fecha": 
                currentCell.alignment = Alignment(horizontal='center')                   
            elif subtitleName == "público g." or subtitleName == "clientes": 
                currentCell.alignment = Alignment(horizontal='right')
                currentCell.number_format = '#,##0.00'
            elif subtitleName == "ing": 
                currentCell.font = Font (size=12, bold=True)
                currentCell.alignment = Alignment(horizontal='center')
                currentCell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type = "solid")
            elif subtitleName == "egr": 
                currentCell.font = Font (size=12, bold=True)
                currentCell.alignment = Alignment(horizontal='center')
                currentCell.fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type = "solid")
            elif subtitleName == "pag": 
                currentCell.font = Font (size=12, bold=True)
                currentCell.alignment = Alignment(horizontal='center')
                currentCell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type = "solid")
            elif titleName == "folio": 
                currentCell.alignment = Alignment(horizontal='center')
            elif titleName == "proveedor": 
                currentCell.alignment = Alignment(horizontal='left')
            elif subtitleName == "importe" or subtitleName == "iva" or titleName == "total":
                currentCell.alignment = Alignment(horizontal='right')
                currentCell.number_format = '#,##0.00'  
            elif titleName == "comentarios": 
                currentCell.alignment = Alignment(horizontal='left')

    wb.save (filePath)






