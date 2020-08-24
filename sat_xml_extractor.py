#! python3
#Extract infomation from xml files and copy to the clipboard

import os, re, pyperclip, logging
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

def extract_information_folder (folder):
    """Extract information from xml files"""
    allFilesInfo = []
    for file in os.listdir(folder):
        fileInfo = {} 

        #Open file and get root
        tree = ElementTree.parse(os.path.join(folder, file)) 
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

        
        allFilesInfo.append(fileInfo)
    return allFilesInfo

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
    """Add format to text"""

    listText = format_specific_gss(data)
    formatedText = ''

    for xmlList in listText: 
        for item in xmlList: 
            formatedText += str(item) + '\t'
        formatedText += '\n'
        
    logging.debug(formatedText)
    pyperclip.copy(formatedText)
    return formatedText

dirname = os.path.dirname(__file__)
pathFolder = os.path.join(dirname, 'XML')
allFilesInfo = extract_information_folder(pathFolder)
text = format_text(allFilesInfo)

"""
def extract_xml (path, patterns):
    #Search an specific regular expreion in each file in a directory
    allMatchesData = []

    for file in os.listdir(path):               #Loop each file in folder
        match_text = {}
        for item in patterns:                   #Loop eacyh pattern 

            documentRegex = re.compile(re.escape(r' ' + item) + r'''(
                ="
                ([a-zA-Z0-9-:.,; ]+)
                (&amp;)?
                ([a-zA-Z0-9-:.,; ]+)?
                " 
            )''', re.VERBOSE) 
            

            text = open(path+'/'+file, 'r').read()
            match = documentRegex.search(text)

            if match:
                if match.group(4):  #Ampersand find in text
                    match_text[item] = (match.group(2) + '&' + match.group(4))
                else:               #Without ampersand
                    match_text[item] = match.group(2)
                
        allMatchesData.append(match_text)

    return allMatchesData
"""



"""
my_data = extract_xml (path, data)
format_text_copy_xml(my_data)
"""