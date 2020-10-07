#! python3
# Extract, format, shorted and process info from xml files

import string, os
from xml.etree import ElementTree


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

def shortXmlByDate (xmlsInfo): 
    """ Short a list of dicctionaries from xml info"""
    dates = []
    shortedData = []

    # Get dates
    for xmlInfo in xmlsInfo: 
        dates.append(xmlInfo['fecha'])
    
    #Short data
    dates.sort()
    for date in dates: 
        for xmlInfo in xmlsInfo: 
            if date == xmlInfo['fecha']: 
                shortedData.append (xmlInfo)
        
    return shortedData

def formatDataEgresos (allInfoXml): 
    """Format data to Egresos info"""
    formatedData = []
    uuids = []

    # Write data in the file
    for info in allInfoXml: 
        currentInfo = []

        # Date
        currentInfo.append (info['fecha'])
        currentInfo.append (info['fecha'][0:10])

        # Type of CFDI
        if info['comprobante'] == 'I':
            currentInfo.append ('X')
            currentInfo.append (" ")
            currentInfo.append (" ")
        elif info['comprobante'] == 'P':
            currentInfo.append (' ')
            currentInfo.append ("X")
            currentInfo.append (" ")
        elif info['comprobante'] == 'E': 
            currentInfo.append (' ')
            currentInfo.append (" ")
            currentInfo.append ("X")

        # UUID
        uuidInUse = False
        for uuid in uuids: 
            if uuid[-4:] == info['uuid'][-4]: 
                # Dont repeat uuids
                uuidInUse = True

        if uuidInUse: 
            currentInfo.append (info['uuid'][-8:])
        else: 
            currentInfo.append (info['uuid'][-4:])
        uuids.append(info['uuid'])
        
        # Name
        currentInfo.append (info['emisor'])

        #quantities
        importe = info['subtotal'] - info ['descuento']
        iva = info ['total'] - importe
        total = info['total']

        if 'metodoPago' in info.keys() and info['metodoPago'] == 'PPD': #Check if is PDD
            currentInfo.append ("")
            currentInfo.append ("")
            currentInfo.append (importe)
            currentInfo.append (iva)
            currentInfo.append (total)
        else:
            currentInfo.append (importe)
            currentInfo.append (iva)
            currentInfo.append ("")
            currentInfo.append ("")
            currentInfo.append (total)


        currentInfo.append ("")
        formatedData.append (currentInfo)

    return formatedData

def writeData (sheet, col, row, data):
    currentRow = row
    chars = list(string.ascii_lowercase)

    for dataLine in data: 
        currentCol = col 
        for dataItem in dataLine:
            sheet[chars[currentCol-1] + str(currentRow)] = dataItem
            currentCol += 1
        currentRow += 1 

def getXmlEgresosInfo (folder):
    """ Get info by xml files, short and format"""
    extractedInfo = []

    # Save data in a list
    for file in os.listdir(os.path.join(folder)):
        if file.endswith('.xml'): 
            xlmPath = os.path.join (folder, file)
            info = extract_information_xml(xlmPath)
            extractedInfo.append (info)
        
    # Short data
    shortedInfo = shortXmlByDate (extractedInfo)
    formatedInfo = formatDataEgresos(shortedInfo)
    return formatedInfo

