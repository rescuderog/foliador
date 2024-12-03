import os
import io
import locale
import re
from comtypes.client import CreateObject
from pypdf import PdfReader, PdfWriter
from docxtpl import DocxTemplate
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from datetime import datetime
from num2words import num2words
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from unidecode import unidecode
from pdf2image import convert_from_path

# metodos principales


def registrar_fuente_custom(font_file):
    pdfmetrics.registerFont(TTFont('Verdana', font_file))


def combinar_cwd_dir(directorio):
    return os.path.join(os.getcwd(), directorio)


def leer_archivos_en_carpeta(carpeta, nomenclatura):
    nomenclatura_match = re.compile(r'(.*)_(.*)_(.*)_(.*)')
    folderPath = combinar_cwd_dir(carpeta)
    hay_img = False
    listArchivos = []
    for i, file in enumerate(os.listdir(folderPath)):
        if os.path.isfile(os.path.join(folderPath, file)):
            nombre, extension = os.path.splitext(file)
            if nombre.startswith("+"):
                nombre = nombre[1:]
                is_img = True
                hay_img = True
            else:
                is_img = False
            if nomenclatura:
                try:
                    nombre = nomenclatura_match.match(nombre).group(4)
                except:
                    pass
            listArchivos.append({
                'nombre': nombre,
                'ext': extension,
                'completo': file,
                'ruta': os.path.join(folderPath, file),
                'is_img': is_img
            })
    listArchivos.sort(key=lambda s: unidecode(s['nombre']).casefold())
    return listArchivos, hay_img


def chequear_word(extension):
    if extension == ".doc" or extension == ".docx":
        return True
    else:
        return False  

def convertir_a_jpeg(pdf_path, dpi=300):
    images = convert_from_path(pdf_path=pdf_path, dpi=dpi, poppler_path=combinar_cwd_dir('./poppler/bin'))
    new_images = []
    for img in images:
        newimg = img.convert('RGB')
        new_images.append(newimg)
    image1 = new_images.pop(0)
    new_img_path = os.path.join(combinar_cwd_dir('./tmpimgs/'), f'TempImg-{datetime.now().timestamp()}.pdf')
    image1.save(fp=new_img_path, save_all=True, append_images=new_images, optimize=True, quality=90)
    return new_img_path


def convertir_a_pdf(archivo_ruta, archivo_nombre, target_dir):
    # recibe ruta completa del archivo, retorna False si no es .doc, retorna el path del archivo
    # si se convierte el archivo sin errores
    if not chequear_word(os.path.splitext(archivo_ruta)[1]):
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
    # crea la pagina solapada de folio, retorna el paquete de bytes
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont("Verdana", 18)
    can.roundRect(485, 730, 100, 50, 1.5)
    can.drawImage("logoUCAlong.jpg", 490, 750, 90, 30,
                  mask='auto',  preserveAspectRatio=True)
    # can.drawString(529, 967, "UCA")
    can.setFont("Verdana", 11)
    can.drawString(489, 740, folioStr)
    can.showPage()
    can.save()
    return packet


def getNumOfPages(pdf_list):
    totalNumOfPages = 0
    for pdf in pdf_list:
        pdf = PdfReader(open(pdf, 'rb'))
        numPages = len(pdf.pages)
        totalNumOfPages = totalNumOfPages + numPages
    return totalNumOfPages


def foliar_archivo(folio_start, archivo_ruta, num_total_pags, dpi, is_img: bool, output: PdfWriter):
    # lee el pdf y agrega la hoja foliada (solapa una hoja arriba de la existente).
    # el output es el nro de folios en el que quedo el contador.
    if is_img:
        archivo_ruta = convertir_a_jpeg(archivo_ruta, dpi)
    existing_pdf = PdfReader(open(archivo_ruta, 'rb'))
    for i in range(0, len(existing_pdf.pages)):
        folio_start = folio_start + 1
        folioStr = f"Folio {folio_start:03d} de {num_total_pags:03d}"
        page = existing_pdf.pages[i]
        h = page.mediabox.height
        w = page.mediabox.width
        orientation = existing_pdf.pages[i].mediabox
        is_landscape = False
        if orientation.right - orientation.left > orientation.top - orientation.bottom:
            page = existing_pdf.pages[i].rotate(270)
            is_landscape = True

        if is_landscape:
            page.transfer_rotation_to_content()

        if h != 1008 or w != 612:
            page.scale_to(width=612, height=1008)

        # create a new PDF with Reportlab
        packet = io.BytesIO()
        packet = createFolioPage(folioStr, packet)
        # move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfReader(packet)
        new_pdf.pages[0].scale_to(width=612, height=1008)
        # add the "watermark" (which is the new pdf) on the existing page
        new_pdf.pages[0].merge_page(page, over=False)
        output.add_page(new_pdf.pages[0])
    return folio_start


def checkGender(genderValue):
    if genderValue == "M":
        return "del Sr."
    elif genderValue == "F":
        return "de la Srta."
    else:
        return " "


def generate_materias(lista_materias):
    materiaList = []
    for i, materia in enumerate(lista_materias):
        materia = f'{i+1}.{materia}'
        if i != 0:
            materia = "\n" + materia
        materiaList.append(materia)
    return materiaList


def generateWordDocUH(context, targetFile, result_dir):
    templateDocx = DocxTemplate(targetFile)
    templateDocx.render(context)
    saveFileName = os.path.join(
        result_dir, f'ResultingUltimaHoja-{datetime.now().timestamp()}.docx')
    templateDocx.save(saveFileName)
    return saveFileName


def set_config_uh(listaDatos, numTotalPages, simulation=False):
    targetFile = listaDatos[8]
    listNombreMaterias = generate_materias(listaDatos[0])
    numTotalWords = num2words(numTotalPages, lang='es')
    locale.setlocale(locale.LC_ALL, '')
    fechaylugar = datetime.today().strftime("Buenos Aires, %d de %B de %Y")
    if simulation:
        context = {'numfolios': 500, 'foliosletras': numTotalWords, 'asigList': listNombreMaterias,
                   'sexo': checkGender(listaDatos[4]), 'apellido': listaDatos[3], 'nombre': listaDatos[2],
                   'dni': listaDatos[1], 'fechaylugar': fechaylugar, 'antequien': listaDatos[5],
                   'carrera': listaDatos[6]}
    else:
        context = {'numfolios': numTotalPages, 'foliosletras': numTotalWords, 'asigList': listNombreMaterias,
                   'sexo': checkGender(listaDatos[4]), 'apellido': listaDatos[3], 'nombre': listaDatos[2],
                   'dni': listaDatos[1], 'fechaylugar': fechaylugar, 'antequien': listaDatos[5],
                   'carrera': listaDatos[6]}

    return targetFile, context


def simulate_generar_uh(listaDatos, numTotalPages, result_dir):
    targetFile, context_test = set_config_uh(
        listaDatos, numTotalPages, simulation=True)
    temp_file = generateWordDocUH(context_test, targetFile, result_dir)
    ultimahojaPDFTemp = PdfReader(open(convertir_a_pdf(
        temp_file, 'ultimaHojaComp', result_dir), 'rb'))
    pages = len(ultimahojaPDFTemp.pages)
    numTotalPages = numTotalPages + pages
    return numTotalPages


def generateUltimaHoja(listaDatos, numTotalPages, result_dir):
    targetFile, context = set_config_uh(listaDatos, numTotalPages)
    saveFileName = generateWordDocUH(context, targetFile, result_dir)
    saveFileName_uhsa = None
    if listaDatos[7]:
        fecha = datetime.today().strftime("%d/%m/%Y")
        context_uhsa = {'fecha': fecha}
        targetFile_uhsa = "modelouhsa.docx"
        templateDocx_uhsa = DocxTemplate(targetFile_uhsa)
        templateDocx_uhsa.render(context_uhsa)
        saveFileName_uhsa = os.path.join(
            result_dir, f'UHSA-{datetime.now().timestamp()}.docx')
        templateDocx_uhsa.save(saveFileName_uhsa)

    return saveFileName, saveFileName_uhsa


def consolidar_pdf(target_dir, result_dir, result_file: PdfWriter, listaDatos, numberOfPages):
    # consolida el output del PdfFileWriter en un pdf, y le agrega la ultima hoja y la UHSA si aplica
    outputFilename = os.path.join(
        result_dir, f"ProgramasCompilados-{datetime.now().timestamp()}.pdf")
    outputStream = open(outputFilename, "wb")
    ultimahoja, ultimahoja_UHSA = generateUltimaHoja(
        listaDatos, numberOfPages, result_dir)
    uhsaPDF = None
    if ultimahoja_UHSA:
        uhsaPDF = PdfReader(open(convertir_a_pdf(
            ultimahoja_UHSA, 'uhsa', target_dir), 'rb'))
    ultimahojaPDF = PdfReader(open(convertir_a_pdf(
        ultimahoja, 'ultimaHojaComp', target_dir), 'rb'))
    for i in range(0, len(ultimahojaPDF.pages)):
        pagenumber = numberOfPages - len(ultimahojaPDF.pages) + i + 1
        folioStr = f"Folio {pagenumber:03d} de {numberOfPages:03d}"
        page = ultimahojaPDF.pages[i]
        page.scale_to(width=612, height=1008)
        packet = io.BytesIO()
        createFolioPage(folioStr, packet)
        packet.seek(0)
        new_pdf = PdfReader(packet)
        new_pdf.pages[0].scale_to(width=612, height=1008)
        page.merge_page(new_pdf.pages[0])
        result_file.add_page(page)

    if uhsaPDF:
        pageUHSA = uhsaPDF.pages[0]
        result_file.add_page(pageUHSA)
    result_file.write(outputStream)
    outputStream.close()
    return True
