'''
 *
 * Python 3.8 || >
 *
 * El objetivo de este script es leer las facturas de AFIP
 * y crear un excel (xlsx) con toda la información fue extraída de las facturas (pdf)
 *
 * @category AFIP
 *
 * @author   MARTIN DELL'ORO 
 *
 * @license  https://www.gnu.org/licenses/gpl-3.0.html GPL v3
 *
 * @link     martindell.oro@gmail.com
'''
import PyPDF2
import os
from time import sleep
import xlsxwriter
import datetime

def isDate (date):
    try:
        datetime.datetime.strptime(date, '%d/%m/%Y')
        return True
    except:
        return False

def isCuit (cuit):
    if len(cuit) == 11:
        return True
    return False

def split(strng, sep, pos):
    strng = strng.split(sep)
    return sep.join(strng[:pos]), sep.join(strng[pos:])

path = os.path.dirname(os.path.realpath(__file__))
facturas = sorted(os.listdir(path))
data = []
pdfSinProcesar = []
workbook = xlsxwriter.Workbook(os.path.join(path, "", 'facturas.xlsx'))
worksheet = workbook.add_worksheet('Hoja 1')
for i in facturas:
    if ("pdf" in i):
        pdfFileObj = open(os.path.join(path, "", i), 'rb')
        pdfReader = PyPDF2.PdfFileReader(os.path.join(path, "", i))
        pdf_writer = PyPDF2.PdfFileWriter()
        if pdfReader.isEncrypted:
            pdfReader.decrypt('')
            pdf_writer.addPage(pdfReader.getPage(0))
            with open(i, "wb") as f:
                pdf_writer.write(f)
        pdfFileObj.close()
row = 2
format = workbook.add_format({'bg_color': '#bcbcbc', 'bold': True})
worksheet.write('A1', 'Razón Social', format)
worksheet.write('B1', 'Cuit', format)
worksheet.write('C1', 'Fecha de emisión', format)
worksheet.write('D1', 'Factura/Nota Crédito', format)
worksheet.write('E1', 'Tipo Letra', format)
worksheet.write('F1', 'Sucursal', format)
worksheet.write('G1', 'Nro. Factura', format)
worksheet.write('H1', 'Cliente', format)
worksheet.write('I1', 'CUIT Cliente', format)
worksheet.write('J1', 'Importe total', format)
worksheet.write('K1', 'Domicilio Cliente', format)
worksheet.write('L1', 'CAE', format)

for i in facturas:
    if ("pdf" in i):
        pdfFileObj = open(os.path.join(path, "", i), 'rb')
        pdfReader = PyPDF2.PdfFileReader(os.path.join(path, "", i), strict=False,)
        pageObj = pdfReader.getPage(0)
        page = pageObj.extractText()
        pageObj = pdfReader.getPage(0)
        page = pageObj.extractText()
        split = "\n"
        spl=page.split(split)
        if ("Fecha de Emisión:" in spl):
            indexAdd = spl.index('Domicilio:')
            cuit = spl[indexAdd + 1] if isDate(spl[indexAdd + 1]) == False else spl[indexAdd + 2] if isDate(spl[indexAdd + 2]) == False else spl[indexAdd + 3] if isDate(spl[indexAdd + 3]) == False else spl[indexAdd + 4] if isDate(spl[indexAdd + 4]) == False else spl[indexAdd + 5] if isDate(spl[indexAdd + 5]) == False else 3
            print(cuit, i)
            tableformat = workbook.add_format({'bg_color': '#bcbcbc'}) if row%2 == 1 else None
            if (isCuit(cuit)):
                razonSocial = spl[3]
                worksheet.write(f'A{row}', razonSocial, tableformat)
                worksheet.write(f'B{row}', cuit, tableformat)
                fechaEmision = spl[spl.index(cuit)-1]
                worksheet.write(f'C{row}', fechaEmision, tableformat)
                condIva = [i for i, n in enumerate(spl) if n == 'Condición frente al IVA:'][1]
                worksheet.write(f'D{row}', spl[condIva+1], tableformat)
                worksheet.write(f'E{row}', spl[condIva+2], tableformat)
                compNro = spl.index('Comp. Nro:')
                worksheet.write(f'F{row}', spl[compNro+1], tableformat)
                worksheet.write(f'G{row}', spl[compNro+2], tableformat)
                cliente = spl[spl.index(cuit)+1]
                worksheet.write(f'H{row}', cliente, tableformat)
                cuitCliente = spl[spl.index('CUIT: ')+1] if 'CUIT: ' in spl else spl[spl.index('DNI: ')+1]
                worksheet.write(f'I{row}', cuitCliente, tableformat)
                total = spl[spl.index('Subtotal: $')-1]
                worksheet.write(f'J{row}', total, tableformat)
                i = spl.index(cuit)+1
                contadoIndex = spl.index('Contado') if 'Contado' in spl else spl.index('Cuenta Corriente') if 'Cuenta Corriente' in spl else spl.index('Otra') if 'Otra' in spl else spl.index('Cheque') if 'Cheque' in spl else spl.index('Tarjeta de Débito') if 'Tarjeta de Débito' in spl else spl.index('Tarjeta de Crédito') if 'Tarjeta de Crédito' in spl else spl.index('Ticket')
                domicilio = ''
                while i < contadoIndex:
                    domicilio = domicilio + ' ' + spl[i]
                    i = i+1
                worksheet.write(f'K{row}', domicilio, tableformat)
                cae = spl.index('Comprobante Autorizado')+3
                worksheet.write(f'L{row}', spl[cae], tableformat)
                row+=1
            else:
                pdfSinProcesar.append([i])
        else:
            pdfSinProcesar.append([i])
        pdfFileObj.close()
workbook.close()

print("\nPDF's que no se han podido procesar:", pdfSinProcesar, '\n')

name=input("Presioná enter para terminar :) -> si tenes alguna pregunta: martindell.oro@gmail.com")
sleep(1)
