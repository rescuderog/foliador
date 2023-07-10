from metodos import leer_archivos_en_carpeta, convertir_a_pdf, combinar_cwd_dir, getNumOfPages, foliar_archivo, consolidar_pdf, registrar_fuente_custom
from pypdf import PdfWriter
import json

#App de consola para generar pdfs foliados en base a lo contenido en una carpeta target
CARPETA_TARGET = './por_foliar/'
CARPETA_RESULTADO = './foliado/'
CARPETA_TEMP = './tmpfolder/'

ultimahoja_fac = None
with open('listFacultades.json', encoding="utf8") as f:
    data = json.load(f)
    print("----------Facultades----------\n")
    for facultad in data['facultades'].keys():
        print(f'{facultad}. {data["facultades"][facultad]["facultad"]}\n')
    facultad_sel = input("Seleccione la facultad deseada ingresando el número: ")
    ultimahoja_fac = data["facultades"][facultad_sel]["hoja_modelo"]

dni = input('Ingresar el DNI del alumno: ')
nombre = input('Ingresar el nombre del alumno: ')
apellido = input('Ingresar el apellido del alumno: ')
sexo = input('Ingresar sexo del alumno (SOLO M o F): ')
antequien = input('SOLO SI LO PIDE, ingresar ante quien se lo presenta: ')
carrera = input('Ingrese la carrera del alumno: ')


diccionario = leer_archivos_en_carpeta(CARPETA_TARGET)
list_archivos = []
list_names = []
list_materias = []

for key in diccionario.keys():
    path = convertir_a_pdf(diccionario[key]['ruta'], diccionario[key]['nombre'], combinar_cwd_dir(CARPETA_TARGET))
    if path:
        list_archivos.append(path)
        list_names.append(diccionario[key]['nombre'])
    else:
        list_archivos.append(diccionario[key]['ruta'])
        list_names.append(diccionario[key]['nombre'])

for i, nombre_archivo in enumerate(list_names):
    new_name = input(f'El nombre del archivo es {nombre_archivo}, si no lo quiere cambiar para la hoja final presione ENTER. De lo contrario, ingrese el nuevo nombre: ')
    if new_name:
        list_materias.append(new_name)
    else:
        list_materias.append(nombre_archivo)

uhsa = input('UHSA? (Ingresar algo, de lo contrario, presionar ENTER): ')
if uhsa:
    uhsa = True
else:
    uhsa = False

datosAlumno = [list_materias, dni, nombre, apellido, sexo, antequien, carrera, uhsa, ultimahoja_fac]

numPags = getNumOfPages(list_archivos)
folio = 0
output = PdfWriter()
registrar_fuente_custom('Verdana.ttf')

for archivo in list_archivos:
    folio_num = foliar_archivo(folio, archivo, numPags, output)
    folio = folio_num

consolidar_pdf(combinar_cwd_dir(CARPETA_TARGET), combinar_cwd_dir(CARPETA_RESULTADO), output, datosAlumno, numPags)