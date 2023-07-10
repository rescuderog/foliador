import os, io, locale
from comtypes.client import CreateObject
from PyPDF2 import PdfReader, PdfWriter
from docxtpl import DocxTemplate, RichText
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from datetime import datetime
from num2words import num2words
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

#metodos principales

def registrar_fuente_custom(font_file):
    pdfmetrics.registerFont(TTFont('Verdana', font_file))

def combinar_cwd_dir(directorio):
    return os.path.join(os.getcwd(), directorio)

def leer_archivos_en_carpeta(carpeta):
    folderPath = combinar_cwd_dir(carpeta)
    dictArchivos = {}
    for i, file in enumerate(os.listdir(folderPath)):
        if os.path.isfile(os.path.join(folderPath, file)):
            nombre, extension = os.path.splitext(file)
            dictArchivos[i] = {
                'nombre': nombre,
                'ext': extension,
                'completo': file,
                'ruta': os.path.join(folderPath, file)
            }
    return dictArchivos

def convertir_a_pdf(archivo_ruta, archivo_nombre, target_dir):
    #recibe ruta completa del archivo, retorna False si no es .doc, retorna el path del archivo
    #si se convierte el archivo sin errores
    if(os.path.splitext(archivo_ruta)[1]) != '.doc':
        return False
    archivo_nombre = archivo_nombre + '.pdf'
    wdToPDF = CreateObject("Word.Application")
    wdFormatPDF = 17
    pdfCreate = wdToPDF.Documents.Open(archivo_ruta)
    resulting_path = os.path.join(target_dir, archivo_nombre)
    pdfCreate.SaveAs(resulting_path, wdFormatPDF)
    pdfCreate.Close()
    wdToPDF.Quit()
    return resulting_path

def createFolioPage(folioStr, packet):
    #crea la pagina solapada de folio, retorna el paquete de bytes
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont("Verdana", 18)
    can.roundRect(485, 940, 100, 50, 1.5)
    can.drawImage("logoUCAlong.jpg", 490, 960, 90, 30, mask='auto',  preserveAspectRatio=True)
    #can.drawString(529, 967, "UCA")
    can.setFont("Verdana", 11)
    can.drawString(489, 950, folioStr)
    can.save()
    return packet

def getNumOfPages(pdf_list):
    totalNumOfPages = 0
    for pdf in pdf_list:
        pdf = PdfReader(open(pdf, 'rb'))
        numPages = len(pdf.pages)
        totalNumOfPages = totalNumOfPages + numPages
    totalNumOfPages = totalNumOfPages + 1
    return totalNumOfPages

def foliar_archivo(folio_start, archivo_ruta, num_total_pags, output: PdfWriter):
    #lee el pdf y agrega la hoja foliada (solapa una hoja arriba de la existente).
    #el output es el nro de folios en el que quedo el contador.
    existing_pdf = PdfReader(open(archivo_ruta, 'rb'))
    for i in range(0, len(existing_pdf.pages)):
        folio_start = folio_start + 1
        folioStr = f"Folio {folio_start:03d} de {num_total_pags:03d}"
        page = existing_pdf.pages[i]
        h = page.mediabox.height
        w = page.mediabox.width
        orientation = existing_pdf.pages[i].mediabox
        if orientation.right - orientation.left > orientation.top - orientation.bottom:
            page = existing_pdf.pages[i].rotateCounterClockwise(90)
        if h != 1008 or w != 612:
            page.scale_to(width=612, height=1008)
        # create a new PDF with Reportlab
        packet = createFolioPage(folioStr, io.BytesIO())
        #move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfReader(packet)
        # add the "watermark" (which is the new pdf) on the existing page
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)
    return folio_start

def checkGender(genderValue):
    if genderValue == "M":
        return "del Sr."
    elif genderValue == "F":
        return "de la Srta."
    else:
        return " "

def generateUltimaHoja(listCodMaterias, listaDatos, numTotalPages, result_dir):
    magicCharNumberPerLine = 110
    targetFile = 'modeloultimahoja.docx'
    templateDocx = DocxTemplate(targetFile)
    numTotalWords = num2words(numTotalPages, lang='es')
    locale.setlocale(locale.LC_ALL, '')
    fechaylugar = datetime.today().strftime("Buenos Aires, %d de %B de %Y")
    fechaylugar = fechaylugar.ljust(magicCharNumberPerLine, '-')
    remainingSpaceEstAsig = 87 - len(str(numTotalPages)) - len(numTotalWords)
    asigLineas = ''.ljust(remainingSpaceEstAsig, '-')
    remainingSpaceEstFinal = 95 - len('Sr.') - len(listaDatos[2]) - len(listaDatos[3]) - len(listaDatos[1])
    finalLineas = ''.ljust(remainingSpaceEstFinal, '-')
    context = {'numfolios': numTotalPages, 'foliosletras': numTotalWords, 'asigList': None, 'sexo': checkGender(listaDatos[4]), 
    'apellido': listaDatos[3], 'nombre': listaDatos[2], 'dni': listaDatos[1], 'fechaylugar': fechaylugar, 
    'asigLineas': asigLineas, 'finalLineas': finalLineas, 'antequien': listaDatos[5]}
    templateDocx.render(context)
    saveFileName = os.path.join(result_dir, f'ResultingUltimaHoja-{datetime.now().timestamp()}.docx')
    templateDocx.save(saveFileName)
    saveFileName_uhsa = None
    if listaDatos[6]:
        fecha = datetime.today().strftime("%d/%m/%Y")
        context_uhsa = {'fecha': fecha}
        targetFile_uhsa = "modelouhsa.docx"
        templateDocx_uhsa = DocxTemplate(targetFile_uhsa)
        templateDocx_uhsa.render(context_uhsa)
        saveFileName_uhsa = os.path.join(result_dir, f'UHSA-{datetime.now().timestamp()}.docx')
        templateDocx_uhsa.save(saveFileName_uhsa)

    return saveFileName, saveFileName_uhsa 

def consolidar_pdf(target_dir, result_dir, result_file: PdfWriter, listaDatos, numberOfPages):
    #consolida el output del PdfFileWriter en un pdf, y le agrega la ultima hoja y la UHSA si aplica
    outputFilename = os.path.join(result_dir, f"ProgramasCompilados-{datetime.now().timestamp()}.pdf")
    outputStream = open(outputFilename, "wb")
    ultimahoja, ultimahoja_UHSA = generateUltimaHoja(None, listaDatos, numberOfPages, result_dir)
    uhsaPDF = None
    if ultimahoja_UHSA:
        uhsaPDF = PdfReader(open(convertir_a_pdf(ultimahoja_UHSA, 'uhsa', target_dir), 'rb'))
    ultimahojaPDF = PdfReader(open(convertir_a_pdf(ultimahoja, 'uhsa', target_dir), 'rb'))
    page = ultimahojaPDF.pages[0]
    w = page.mediabox.height
    h = page.mediabox.width
    pagenumber = numberOfPages
    folioStr = f"Folio {pagenumber:03d} de {numberOfPages:03d}"
    packet = io.BytesIO()
    createFolioPage(folioStr, packet)
    packet.seek(0)
    new_pdf = PdfReader(packet)
    page.merge_page(new_pdf.pages[0])
    result_file.add_page(page)
    if uhsaPDF:
        pageUHSA = uhsaPDF.pages[0]
        result_file.add_page(pageUHSA)
    result_file.write(outputStream)
    outputStream.close()
    return True