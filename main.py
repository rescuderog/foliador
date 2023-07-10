from metodos import leer_archivos_en_carpeta, convertir_a_pdf, combinar_cwd_dir, getNumOfPages, foliar_archivo, consolidar_pdf, registrar_fuente_custom
from PyPDF2 import PdfWriter

#App de consola para generar pdfs foliados en base a lo contenido en una carpeta target
CARPETA_TARGET = './por_foliar/'
CARPETA_RESULTADO = './foliado/'
CARPETA_TEMP = './tmpfolder/'

dni = input('Ingresar el DNI del alumno: ')
nombre = input('Ingresar el nombre del alumno: ')
apellido = input('Ingresar el apellido del alumno: ')
sexo = input('Ingresar sexo del alumno (SOLO M o F): ')
antequien = input('SOLO SI LO PIDE, ingresar ante quien se lo presenta: ')

datosAlumno = ['padding', dni, nombre, apellido, sexo, antequien, False]

diccionario = leer_archivos_en_carpeta(CARPETA_TARGET)
list_archivos = []
for key in diccionario.keys():
    path = convertir_a_pdf(diccionario[key]['ruta'], diccionario[key]['nombre'], combinar_cwd_dir(CARPETA_TARGET))
    if path:
        list_archivos.append(path)
    else:
        list_archivos.append(diccionario[key]['ruta'])

numPags = getNumOfPages(list_archivos)
folio = 0
output = PdfWriter()
registrar_fuente_custom('Verdana.ttf')

for archivo in list_archivos:
    folio_num = foliar_archivo(folio, archivo, numPags, output)
    folio = folio_num

consolidar_pdf(combinar_cwd_dir(CARPETA_TARGET), combinar_cwd_dir(CARPETA_RESULTADO), output, datosAlumno, numPags)