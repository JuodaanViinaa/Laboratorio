import json
from Funciones import *

archivo = 'Prueba.xlsx'
# directorioBrutos = 'C:/Users/Admin/Desktop/Escape/Datos/Brutos/'
# directorioConvertidos = 'C:/Users/Admin/Desktop/Escape/Datos/ConvertidosPython/Escape/'
directorioTemporal = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Temporal/'
directorioBrutos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/PruebasBrutos/'
directorioConvertidos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Flex/'

sujetos = ['E3', 'E4', 'E5', 'E7', 'E8', 'E9', 'E1']
columnasProp = [2, 3, 4, 5, 6, 7, 8]
columnasResp = [2, 9, 16, 23, 30, 37, 44]
columnasLatPal = [2, 7, 12, 17, 22, 27, 32]
columnasEscapes = [2, 13, 24, 35, 46, 57, 68]
columnasLatEsc = [2, 13, 24, 35, 46, 57, 68]
columnasEscForz = [2, 7, 12, 17, 22, 27, 32]
sesionesPresentes = []  # Lista vacía.
marcadores = []  # Lista vacía.
tiempo = []  # Lista vacía.
#
# analysis_list = [
#     # Respuestas palancas
#     # Respuestas Pal forzados discriminativos reforzados
#     {"conteoresp": {"mark1": 114, "mark2": 180, "mark3": 202,
#                     "label": "PalDiscRef",
#                     "sheet_position": 1,
#                     "column": 1,
#                     "substract": True,
#                     "summary_column_list": columnasResp,
#                     "offset": 0}},
#     # Respuestas Pal forzados discriminativos no reforzados
#     {"conteoresp": {"mark1": 115, "mark2": 180, "mark3": 202,
#                     "label": "PalNoDiscRef",
#                     "sheet_position": 1,
#                     "column": 3,
#                     "substract": True,
#                     "summary_column_list": columnasResp,
#                     "offset": 1}},
#     # Respuestas Pal forzados no discriminativos 1
#     {"conteoresp": {"mark1": 134, "mark2": 180, "mark3": 201,
#                     "label": "PalNoDisc1",
#                     "sheet_position": 1,
#                     "column": 5,
#                     "substract": True,
#                     "summary_column_list": columnasResp,
#                     "offset": 2}},
#     # Respuestas Pal forzados no discriminativos 2
#     {"conteoresp": {"mark1": 137, "mark2": 180, "mark3": 201,
#                     "label": "PalNoDisc2",
#                     "sheet_position": 1,
#                     "column": 7,
#                     "substract": True,
#                     "summary_column_list": columnasResp,
#                     "offset": 3}},
#
#     # Latencias palancas
#     # Latencias pal forzados discriminativos reforzados
#     {"conteolat": {"mark1": 112, "mark2": 113,
#                    "label": "LatPalDisc",
#                    "sheet_position": 2,
#                    "column": 9,
#                    "substract": False,
#                    "summary_column_list": columnasLatPal,
#                    "offset": 0}},
#     # Latencias pal forzados no discriminativos
#     {"conteolat": {"mark1": 132, "mark2": 133,
#                    "label": "LatPalNoDisc",
#                    "sheet_position": 2,
#                    "column": 11,
#                    "substract": False,
#                    "summary_column_list": columnasLatPal,
#                    "offset": 1}},
#
#     # Respuestas comederos
#     # Respuestas comederos forzados discriminativos reforzados
#     {"conteoresp": {"mark1": 114, "mark2": 16, "mark3": 203,
#                     "label": "ComDiscRef",
#                     "sheet_position": 3,
#                     "column": 13,
#                     "substract": False,
#                     "summary_column_list": columnasResp,
#                     "offset": 0}},
#     # Respuestas comederos forzados discriminativos no reforzados
#     {"conteoresp": {"mark1": 115, "mark2": 117, "mark3": 203,
#                     "label": "ComNoDiscRef",
#                     "sheet_position": 3,
#                     "column": 15,
#                     "substract": False,
#                     "summary_column_list": columnasResp,
#                     "offset": 1}},
#     # Respuestas comederos forzados no discriminativos 1
#     {"conteoresp": {"mark1": 134, "mark2": 40, "mark3": 203,
#                     "label": "ComNoDisc1",
#                     "sheet_position": 3,
#                     "column": 17,
#                     "substract": False,
#                     "summary_column_list": columnasResp,
#                     "offset": 2}},
#     # Respuestas comederos forzados no discriminativos 2
#     {"conteoresp": {"mark1": 137, "mark2": 43, "mark3": 203,
#                     "label": "ComNoDisc2",
#                     "sheet_position": 3,
#                     "column": 19,
#                     "substract": False,
#                     "summary_column_list": columnasResp,
#                     "offset": 3}},
#
#     # Escapes
#     # Respuestas nosepoke forzados discriminativos reforzados
#     {"conteototal": {"mark1": 301,
#                      "label": "EscForzDiscRef",
#                      "sheet_position": 4,
#                      "column": 21,
#                      "substract": False,
#                      "summary_column_list": columnasEscapes,
#                      "offset": 0}},
#     # Respuestas nosepoke forzados discriminativos no reforzados
#     {"conteototal": {"mark1": 302,
#                      "label": "EscForzDiscNoRef",
#                      "sheet_position": 4,
#                      "column": 23,
#                      "substract": False,
#                      "summary_column_list": columnasEscapes,
#                      "offset": 1}},
#     # Respuestas nosepoke forzados no discriminativos 1
#     {"conteototal": {"mark1": 303,
#                      "label": "EscForzNoDisc1",
#                      "sheet_position": 4,
#                      "column": 25,
#                      "substract": False,
#                      "summary_column_list": columnasEscapes,
#                      "offset": 2}},
#     # Respuestas nosepoke forzados no discriminativos 2
#     {"conteototal": {"mark1": 304,
#                      "label": "EscForzNoDisc2",
#                      "sheet_position": 4,
#                      "column": 27,
#                      "substract": False,
#                      "summary_column_list": columnasEscapes,
#                      "offset": 3}},
#     # Respuestas nosepoke forzados discriminativos reforzados
#     {"conteototal": {"mark1": 305,
#                      "label": "EscLibDiscRef",
#                      "sheet_position": 4,
#                      "column": 29,
#                      "substract": False,
#                      "summary_column_list": columnasEscapes,
#                      "offset": 4}},
#     # Respuestas nosepoke forzados discriminativos no reforzados
#     {"conteototal": {"mark1": 306,
#                      "label": "EscLibDiscNoRef",
#                      "sheet_position": 4,
#                      "column": 31,
#                      "substract": False,
#                      "summary_column_list": columnasEscapes,
#                      "offset": 5}},
#     # Respuestas nosepoke forzados no discriminativos 1
#     {"conteototal": {"mark1": 307,
#                      "label": "EscLibNoDisc1",
#                      "sheet_position": 4,
#                      "column": 33,
#                      "substract": False,
#                      "summary_column_list": columnasEscapes,
#                      "offset": 6}},
#     # Respuestas nosepoke forzados no discriminativos 2
#     {"conteototal": {"mark1": 308,
#                      "label": "EscLibNoDisc2",
#                      "sheet_position": 4,
#                      "column": 35,
#                      "substract": False,
#                      "summary_column_list": columnasEscapes,
#                      "offset": 7}},
#
#     # Latencias Escape
#     # Latencias nosepoke forzados discriminativos reforzados
#     {"conteolat": {"mark1": 114, "mark2": 301,
#                    "label": "LatEscFDiscRef",
#                    "sheet_position": 5,
#                    "column": 37,
#                    "substract": False,
#                    "summary_column_list": columnasEscapes,
#                    "offset": 0}},
#     # Latencias nosepoke forzados discriminativos no reforzados
#     {"conteolat": {"mark1": 115, "mark2": 302,
#                    "label": "LatEscFDiscNoRef",
#                    "sheet_position": 5,
#                    "column": 39,
#                    "substract": False,
#                    "summary_column_list": columnasEscapes,
#                    "offset": 1}},
#     # Latencias nosepoke forzados no discriminativos 1
#     {"conteolat": {"mark1": 134, "mark2": 303,
#                    "label": "LatEscFNoDisc1",
#                    "sheet_position": 5,
#                    "column": 41,
#                    "substract": False,
#                    "summary_column_list": columnasEscapes,
#                    "offset": 2}},
#     # Latencias nosepoke forzados no discriminativos 2
#     {"conteolat": {"mark1": 137, "mark2": 304,
#                    "label": "LatEscFNoDisc2",
#                    "sheet_position": 5,
#                    "column": 43,
#                    "substract": False,
#                    "summary_column_list": columnasEscapes,
#                    "offset": 3}},
#     # Latencias nosepoke libres discriminativos reforzados
#     {"conteolat": {"mark1": 154, "mark2": 305,
#                    "label": "LatEscLDiscRef",
#                    "sheet_position": 5,
#                    "column": 45,
#                    "substract": False,
#                    "summary_column_list": columnasEscapes,
#                    "offset": 4}},
#     # Latencias nosepoke libres discriminativos reforzados
#     {"conteolat": {"mark1": 155, "mark2": 306,
#                    "label": "LatEscLDiscRef",
#                    "sheet_position": 5,
#                    "column": 47,
#                    "substract": False,
#                    "summary_column_list": columnasEscapes,
#                    "offset": 5}},
#     # Latencias nosepoke libres discriminativos reforzados
#     {"conteolat": {"mark1": 157, "mark2": 307,
#                    "label": "LatEscLNoDisc1",
#                    "sheet_position": 5,
#                    "column": 49,
#                    "substract": False,
#                    "summary_column_list": columnasEscapes,
#                    "offset": 6}},
#     # Latencias nosepoke libres discriminativos reforzados
#     {"conteolat": {"mark1": 160, "mark2": 308,
#                    "label": "LatEscLDisc2",
#                    "sheet_position": 5,
#                    "column": 51,
#                    "substract": False,
#                    "summary_column_list": columnasEscapes,
#                    "offset": 7}},
# ]

with open("data.json", "r") as data_file:
    json_list = json.load(data_file)

purgeSessions(directorioTemporal, sujetos, sesionesPresentes, columnasProp, columnasResp, columnasLatPal,
              columnasEscapes, columnasLatEsc, columnasEscForz)
convertir(directorioTemporal, directorioBrutos, directorioConvertidos, sujetos, sesionesPresentes, subfijo="_SUBCHOIL_")

wb = createDocument(archivo, directorioConvertidos)

sheet_list = create_sheets(wb, 'Proporciones', 'Respuestas', 'Latencias', 'Comedero', 'Escapes', 'LatNosepoke',
                           'EscapesForzados', 'LatEscapeForz')

analyze(dirConv=directorioConvertidos, fileName=archivo, subList=sujetos, sessionList=sesionesPresentes,
        suffix="_SUBCHOIL_", workbook=wb, sheetList=sheet_list, analysisList=json_list, markColumn="P",
        timeColumn="O")

wb.save(directorioConvertidos + archivo)
