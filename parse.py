import PyPDF2
import sys
import os
from numberparse import *
import pandas as pd
import openpyxl

excelLines = []

def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start


def parseNumberTarjeta(text):
    number = parseNumber(text)
    if text[-1] == '-':
        return number * -1
    return number

def parsePDFTarjeta(file):

    def PrintsumAndAppend(title,value):
        print(title)
        value = sum(value)
        print(value)
        excelLine.append(value)


    # Abre el archivo PDF y crea un objeto "lector"
    pdf_file = open(file, 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    print("ARCHIVO " + file)

    ventasTotales = []
    aranceles = []
    ivaCF = []
    iibb = []
    sirtac = []
    dtoCuotas = []
    perciva = []
    percgcias = []
    retsirtac = []
    promoahora = []
    iva105 = []
    ajustesirtac = []
    cobranzaTotal = []
    totalpresentado = []
    totalnetopagos = []



    # Itera sobre cada página del archivo
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text = page.extract_text()
        lines = text.split('\n')
        excelLine = []
        excelLine.append(file)
        # Itera sobre cada línea del texto
        for line in lines:

            if 'VENTAS C/DESCUENTO CONTADO' in line:

                ventasTotales.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'VENTAS C/DTO CUOTAS FINANC. OTORG.' in line:

                ventasTotales.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'ARANCEL' in line:

                aranceles.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'CARGO TERMINAL FISERV' in line:

                aranceles.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'CARGO LIQUIDACION' in line:

                aranceles.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'SISTEMA CUOTAS' in line:

                aranceles.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'DESCUENTO FINANC.OTORG. CUOTAS' in line:

                dtoCuotas.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'IVA RI. CARGO' in line:

                ivaCF.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'IVA CRED.FISC.COM.L.25063 S/DTO F.OTOR 10,50%' in line:

                iva105.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'IVA RI SIST' in line:

                ivaCF.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'IVA RI CRED' in line:

                ivaCF.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'IVA CRED.FISC.COMERCIO S/ARANC 21,00%' in line:

                ivaCF.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'ING.BRUTOS   CAPITAL FEDERAL' in line:

                iibb.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'RETENCION ING.BRUTOS SIRTAC' in line:

                sirtac.append(parseNumberTarjeta(line[line.find('$')+2:]))
            elif 'IVA PROMO CUOTAS' in line:

                iva105.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'PROMO CUOTAS AHORA 12/18' in line:

                promoahora.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'PERCEPCION IVA R.G. 2408' in line:

                perciva.append(parseNumberTarjeta(line[line.find('$')+2:]))
            elif 'AJUSTE SIRTAC' in line:

                ajustesirtac.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'IMPORTE NETO DE PAGOS' in line:

                cobranzaTotal.append(parseNumberTarjeta(line[line.find('$')+2:]))

            elif 'TOTAL LIQ. TARJ. CREDITO' in line:
                totalnetopagos.append(parseNumberTarjeta(line[find_nth(line,':',2)+2:]))

            elif 'TOTAL LIQ. TARJ. DEBITO' in line:
                totalnetopagos.append(parseNumberTarjeta(line[find_nth(line,':',2)+2:]))

    PrintsumAndAppend("VENTAS TOTALES",ventasTotales)
    PrintsumAndAppend("ARANCELES",aranceles)
    PrintsumAndAppend("IVA 21%",ivaCF)
    PrintsumAndAppend("IIBB",iibb)
    PrintsumAndAppend("SIRTAC",sirtac)
    PrintsumAndAppend("ARANCEL PROMO AHORA",promoahora)
    PrintsumAndAppend("IVA 10,5%",iva105)
    PrintsumAndAppend("PERCEPCION IVA",perciva)
    PrintsumAndAppend("AJUSTE SIRTAC",ajustesirtac)
    PrintsumAndAppend("DESCUETO CUOTAS",dtoCuotas)
    PrintsumAndAppend("COBRANZAS TOTALES",cobranzaTotal)
    PrintsumAndAppend("TOTAL NETO PAGOS",totalnetopagos)

    print("CONTROL")
    control = sum(ventasTotales) - sum(aranceles) - sum(ivaCF) - sum(iibb) - sum(sirtac) - sum(promoahora) - sum(iva105) - sum(perciva) - sum(dtoCuotas) + sum(ajustesirtac)
    dif = sum(totalnetopagos)-control
    print(dif)
    excelLine.append(dif)
    print("---------------------------------")
    excelLines.append(excelLine)


if len(sys.argv) > 1:
    pathDir = sys.argv[1];

for file in os.listdir(pathDir):
    if file.endswith("pdf"):
        parsePDFTarjeta(pathDir+"/"+file)
df = pd.DataFrame(excelLines, columns=['ARCHIVO','VENTAS TOTALES', 'ARANCELES', 'IVA 21%','IIBB','SIRTAC','ARANCEL PROMO AHORA','IVA 10,5%','PERCEPCION IVA','AJUSTE SIRTAC','DESCUENTO CUOTAS','COBRANZAS TOTALES','TOTAL NETO PAGOS', 'CONTROL'])
df.to_excel(pathDir+'/resumen.xlsx')
