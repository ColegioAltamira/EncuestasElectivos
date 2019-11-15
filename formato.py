import re
import time

import xlwt
import gspread as gs
from oauth2client.service_account import ServiceAccountCredentials

# Conseguimos las credenciales para usar la API de Google Drive
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
cliente = gs.authorize(creds)

# Abrimos las hojas y archivos
encuestas = [cliente.open("Encuesta Electivos I° (respuestas)").sheet1,
             cliente.open("Encuesta Electivos II° (respuestas)").sheet1,
             cliente.open("Encuesta Electivos III° (respuestas)").sheet1]

index_col_cursos = 2

# Creamos el archivo de Excel
workbook = xlwt.Workbook()
worksheets = [workbook.add_sheet('Iº'), workbook.add_sheet('IIº'), workbook.add_sheet('IIIº')]
worksheets_cursores = [0, 0, 0]

for index_encuesta in range(len(encuestas)):
    encuesta = encuestas[index_encuesta]

    col_index = 0

    # Conseguimos la fila de los títulos y la columna de cursos
    titulos = encuesta.row_values(1)
    cursos = encuesta.col_values(index_col_cursos)[1:]

    for col in titulos:
        col_index += 1

        # Intentamos encotrar "[ELECTIVO]"
        electivo = re.findall("\[(.*?)\]", col)

        if len(electivo) == 0:
            # No es un electivo
            continue

        # Si es un electivo entnces escribimos el título
        hoja = worksheets[index_encuesta]
        hoja.write(worksheets_cursores[index_encuesta], 0, electivo[0])
        hoja.write(worksheets_cursores[index_encuesta] + 1, 0, "Curso")
        hoja.write(worksheets_cursores[index_encuesta] + 1, 1, "Nota")
        worksheets_cursores[index_encuesta] += 2

        # Buscamos la columna de las notas correspondientes a este electivo
        notas_electivo = encuesta.col_values(col_index)[1:]

        # Por cada estudiante escribimos el curso y la nota
        for x in range(len(notas_electivo)):
            hoja.write(worksheets_cursores[index_encuesta], 0, cursos[x])
            hoja.write(worksheets_cursores[index_encuesta], 1, notas_electivo[x])
            worksheets_cursores[index_encuesta] += 1

        # Agregamos dos espacios en blanco para que se vea más ordenado
        worksheets_cursores[index_encuesta] += 2

    time.sleep(100)

workbook.save('encuestas.xls')
