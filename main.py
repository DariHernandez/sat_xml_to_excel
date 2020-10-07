#! python3
#Extract infomation from xml files and copy to the clipboard

import os, openpyxl, logging
from xml.etree import ElementTree
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

def getIngresoEgreso (comprobante, cantidad): 
    """Get in negative o positive number in function of type of bill"""
    #Check if the value exist
    if cantidad:
        if comprobante == 'I': 
            # Return positive
            return float(cantidad)
        elif comprobante == 'E':
            #Return negative
            return -float(cantidad)
    else: 
        return 0

def extract_information_xml (file):
    """Extract information from xml files"""
    fileInfo = {} 

    #Open file and get root
    tree = ElementTree.parse(os.path.join(file)) 
    root = tree.getroot()

    #Open specific tasg to get UUID and MONTO
    prefix = '{' + 'http://www.sat.gob.mx/'
    emisor = root.find(prefix + 'cfd/3}Emisor')
    complemento = root.find(prefix + 'cfd/3}Complemento')
    timbreFiscalDigital = complemento.find (prefix + 'TimbreFiscalDigital}TimbreFiscalDigital')
    
    #Extract general information
    fileInfo['fecha']        = root.get('Fecha')
    fileInfo['comprobante']  = root.get('TipoDeComprobante')
    fileInfo['emisor']       = emisor.get('Nombre')
    fileInfo['uuid']         = timbreFiscalDigital.get('UUID')

    # Extract specific information
    if fileInfo['comprobante'] == 'P': 
        # Extract information of type P
        pagos = complemento.find(prefix + 'Pagos}Pagos')
        pago = pagos.find(prefix + 'Pagos}Pago' )
        monto = pago.get('Monto')

        fileInfo['subtotal'] = float(monto)/1.16
        fileInfo['total'] = float(monto)
        fileInfo['descuento'] =  0
    else: 
        # Extract information for type I and E
        fileInfo['metodoPago'] = root.get('MetodoPago')
        fileInfo['subtotal'] = getIngresoEgreso (fileInfo['comprobante'], root.get('SubTotal')) 
        fileInfo['total'] = getIngresoEgreso (fileInfo['comprobante'], root.get('Total')) 
        fileInfo['descuento'] = getIngresoEgreso (fileInfo['comprobante'], root.get('Descuento')) 

    return fileInfo

"""
def format_specific_gss (data):
    # Get format to text getted by xml files, and copy to clipboard
    # Formted to spreadsheed
    
    text = []
    for line in data: 
        textLine = []

        #Fecha 
        textLine.append(line['fecha'])

        #Fecha shorted
        textLine.append(line['fecha'][0:10])

        #Type of invoice: ingreso, egreso, pago
        comprobante = ''
        if line['comprobante'] == 'I':
            comprobante = 'X\t\'\t\''   #white space - whide cell
        elif line['comprobante'] == 'P':
            comprobante = '\'\tX\t\''
        elif line['comprobante'] == 'E': 
            comprobante = '\'\t\'\tX'
        
        textLine.append(comprobante)

        #Folio

        folio = ''
        if data.index(line) > 1:
            for i in range(data.index(line)):
                if line['uuid'] == data[i]['uuid']: 
                    folio = line['uuid'][-8:]
                else: 
                    folio = line['uuid'][-4:]
        else: 
           folio = line['uuid'][-4:]
        
        textLine.append(folio)

        #Emisor
        textLine.append(line['emisor'])

        #Process Importe
        importe = ''
        importe = str(line['subtotal'] - line['descuento']) 

        #Process Iva 
        iva = ''
        iva = str(line['total'] - float(importe)) 
        
        #Process Total
        total = line['total']
        
        #Tipo de pago y montos
        if 'metodoPago' in line.keys() and line['metodoPago'] == 'PPD': #Check if is PDD
            textLine.append('\'')
            textLine.append('\'')
            textLine.append(importe)
            textLine.append(iva)
            textLine.append(total)
        else:  #Then is PEU
            textLine.append(importe)
            textLine.append(iva)
            textLine.append('\'')
            textLine.append('\'')
            textLine.append(total)
        
        text.append(textLine)
    return text

def format_text (data):
    "Add format to text"

    listText = format_specific_gss(data)
    formatedText = ''

    for xmlList in listText: 
        for item in xmlList: 
            formatedText += str(item) + '\t'
        formatedText += '\n'
        
    logging.debug(formatedText)
    pyperclip.copy(formatedText)
    return formatedText
"""

path = "/home/dari/Documentos/dari_developer_fact"

for folder in os.listdir (path):
    if os.path.isdir (folder):  
        filePath = os.path.join (path, (os.path.basename (path) + ".xlsx"))
        print (filePath)

        # Open file o make new file
        try: 
            wb = openpyxl.load_workbook(filePath)
        except FileNotFoundError: 
            wb = openpyxl.Workbook()

        # Make new sheet
        sheet = wb.create_sheet(folder)
        
        # Set info
        uuids = []
        row = 1
        allFilesInfo = []

        # Open and copy info for each xml file
        for file in os.listdir(os.path.join(path, folder)):
            if file.endswith('.xml'): 
                xlmPath = os.path.join (path, folder, file)
                info = extract_information_xml(xlmPath)

                # Date
                sheet['A' + str(row)] = info['fecha']
                sheet['B' + str(row)] = info['fecha'][0:10]

                # Type of CFDI
                if info['comprobante'] == 'I':
                    sheet['C' + str(row)] = 'X'
                elif info['comprobante'] == 'P':
                    sheet['D' + str(row)] = 'X'
                elif info['comprobante'] == 'E': 
                    sheet['E' + str(row)] = 'X'

                # UUID
                ## Dont repeat uuids
                for uuid in uuids: 
                    if uuid[-4:] == info['uuid'][-4]: 
                        sheet['F' + str(row)] = info['fecha'][-8:]
                    else: 
                        sheet['F' + str(row)] = info['fecha'][-4:]
                uuids.append(info['uuid'])
                
                # Name
                sheet['F' + str(row)] = info['emisor']

                #quantities
                importe = info['subtotal'] - info ['descuento']
                iva = info ['total'] - importe
                total = info['total']

                if 'metodoPago' in info.keys() and info['metodoPago'] == 'PPD': #Check if is PDD
                    sheet['G' + str(row)] = ""
                    sheet['H' + str(row)] = ""
                    sheet['I' + str(row)] = importe
                    sheet['J' + str(row)] = iva
                    sheet['K' + str(row)] = total
                else:
                    sheet['G' + str(row)] = importe
                    sheet['H' + str(row)] = iva
                    sheet['I' + str(row)] = ""
                    sheet['J' + str(row)] = ""
                    sheet['K' + str(row)] = total

                sheet['L' + str(row)] = "" 
                row += 1
        wb.save (filePath)

    

"""
allFilesInfo = extract_information_folder(pathFolder)
text = format_text(allFilesInfo)
"""