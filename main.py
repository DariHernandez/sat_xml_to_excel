#! python3
#Extract infomation from xml files and copy to the clipboard

import os, openpyxl, string
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

path = "/home/dari/Documentos/dari_developer_fact"
allInfo = []

for folder in os.listdir (path):
    if os.path.isdir (os.path.join(path, folder)):  
        filePath = os.path.join (path, (os.path.basename (path) + ".xlsx"))

        # Open file o make new file
        try: 
            wb = openpyxl.load_workbook(filePath)
        except FileNotFoundError: 
            wb = openpyxl.Workbook()

        # Make new sheet
        sheetName = folder
        if not sheetName in wb.sheetnames: 
            sheet = wb.create_sheet(folder)
        else: 
            sheet = wb[sheetName]
        
        formatedShortedInfo = getXmlEgresosInfo (os.path.join(path, folder))

        writeInfo (sheet, 1, 1, formatedShortedInfo)

        wb.save (filePath)
