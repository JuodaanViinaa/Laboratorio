import json
from Funciones import *

archivo = 'Prueba2.xlsx'
# directorioBrutos = 'C:/Users/Admin/Desktop/Escape/Datos/Brutos/'
# directorioConvertidos = 'C:/Users/Admin/Desktop/Escape/Datos/ConvertidosPython/Escape/'
directorioTemporal = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Temporal/'
directorioBrutos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/PruebasBrutos/'
directorioConvertidos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Flex/'

sujetos = ['E3', 'E4', 'E5', 'E7', 'E8', 'E9', 'E1']
columnasProp = [2, 4, 6, 8, 10, 12, 14]
columnasResp = [2, 9, 16, 23, 30, 37, 44]
columnasLatPal = [2, 7, 12, 17, 22, 27, 32]
columnasEscapes = [2, 13, 24, 35, 46, 57, 68]
columnasLatEsc = [2, 13, 24, 35, 46, 57, 68]
columnasEscForz = [2, 7, 12, 17, 22, 27, 32]
sesionesPresentes = []  # Lista vacía.
marcadores = []  # Lista vacía.
tiempo = []  # Lista vacía.

analysis_list = [
    # Proporciones
    {"fetch": {"sheet": "Proporciones",
               "summary_column_list": columnasProp,
               "cell_row": 14,
               "cell_column": 6,
               "offset": 0
               }},
    # Ensayos completados
    {"fetch": {"sheet": "Proporciones",
               "summary_column_list": columnasProp,
               "cell_row": 15,
               "cell_column": 2,
               "offset": 1
               }},

    # Respuestas palancas
    # Respuestas Pal forzados discriminativos reforzados
    {"conteoresp": {"mark1": 114, "mark2": 180, "mark3": 202,
                    "label": "PalDiscRef",
                    "sheet": "Respuestas",
                    "column": 1,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 0}},
    # Respuestas Pal forzados discriminativos no reforzados
    {"conteoresp": {"mark1": 115, "mark2": 180, "mark3": 202,
                    "label": "PalNoDiscRef",
                    "sheet": "Respuestas",
                    "column": 3,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 1}},
    # Respuestas Pal forzados no discriminativos 1
    {"conteoresp": {"mark1": 134, "mark2": 180, "mark3": 201,
                    "label": "PalNoDisc1",
                    "sheet": "Respuestas",
                    "column": 5,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 2}},
    # Respuestas Pal forzados no discriminativos 2
    {"conteoresp": {"mark1": 137, "mark2": 180, "mark3": 201,
                    "label": "PalNoDisc2",
                    "sheet": "Respuestas",
                    "column": 7,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 3}},

    # Latencias palancas
    # Latencias pal forzados discriminativos reforzados
    {"conteolat": {"mark1": 112, "mark2": 113,
                   "label": "LatPalDisc",
                   "sheet": "Latencias",
                   "column": 9,
                   "substract": False,
                   "summary_column_list": columnasLatPal,
                   "offset": 0}},
    # Latencias pal forzados no discriminativos
    {"conteolat": {"mark1": 132, "mark2": 133,
                   "label": "LatPalNoDisc",
                   "sheet": "Latencias",
                   "column": 11,
                   "substract": False,
                   "summary_column_list": columnasLatPal,
                   "offset": 1}},

    # Respuestas comederos
    # Respuestas comederos forzados discriminativos reforzados
    {"conteoresp": {"mark1": 114, "mark2": 16, "mark3": 203,
                    "label": "ComDiscRef",
                    "sheet": "Comedero",
                    "column": 13,
                    "substract": False,
                    "summary_column_list": columnasResp,
                    "offset": 0}},
    # Respuestas comederos forzados discriminativos no reforzados
    {"conteoresp": {"mark1": 115, "mark2": 117, "mark3": 203,
                    "label": "ComNoDiscRef",
                    "sheet": "Comedero",
                    "column": 15,
                    "substract": False,
                    "summary_column_list": columnasResp,
                    "offset": 1}},
    # Respuestas comederos forzados no discriminativos 1
    {"conteoresp": {"mark1": 134, "mark2": 40, "mark3": 203,
                    "label": "ComNoDisc1",
                    "sheet": "Comedero",
                    "column": 17,
                    "substract": False,
                    "summary_column_list": columnasResp,
                    "offset": 2}},
    # Respuestas comederos forzados no discriminativos 2
    {"conteoresp": {"mark1": 137, "mark2": 43, "mark3": 203,
                    "label": "ComNoDisc2",
                    "sheet": "Comedero",
                    "column": 19,
                    "substract": False,
                    "summary_column_list": columnasResp,
                    "offset": 3}},

    # Escapes
    # Respuestas nosepoke forzados discriminativos reforzados
    {"conteototal": {"mark1": 301,
                     "label": "EscForzDiscRef",
                     "sheet": "Escapes",
                     "column": 21,
                     "substract": False,
                     "summary_column_list": columnasEscapes,
                     "offset": 0}},
    # Respuestas nosepoke forzados discriminativos no reforzados
    {"conteototal": {"mark1": 302,
                     "label": "EscForzDiscNoRef",
                     "sheet": "Escapes",
                     "column": 23,
                     "substract": False,
                     "summary_column_list": columnasEscapes,
                     "offset": 1}},
    # Respuestas nosepoke forzados no discriminativos 1
    {"conteototal": {"mark1": 303,
                     "label": "EscForzNoDisc1",
                     "sheet": "Escapes",
                     "column": 25,
                     "substract": False,
                     "summary_column_list": columnasEscapes,
                     "offset": 2}},
    # Respuestas nosepoke forzados no discriminativos 2
    {"conteototal": {"mark1": 304,
                     "label": "EscForzNoDisc2",
                     "sheet": "Escapes",
                     "column": 27,
                     "substract": False,
                     "summary_column_list": columnasEscapes,
                     "offset": 3}},
    # Respuestas nosepoke forzados discriminativos reforzados
    {"conteototal": {"mark1": 305,
                     "label": "EscLibDiscRef",
                     "sheet": "Escapes",
                     "column": 29,
                     "substract": False,
                     "summary_column_list": columnasEscapes,
                     "offset": 4}},
    # Respuestas nosepoke forzados discriminativos no reforzados
    {"conteototal": {"mark1": 306,
                     "label": "EscLibDiscNoRef",
                     "sheet": "Escapes",
                     "column": 31,
                     "substract": False,
                     "summary_column_list": columnasEscapes,
                     "offset": 5}},
    # Respuestas nosepoke forzados no discriminativos 1
    {"conteototal": {"mark1": 307,
                     "label": "EscLibNoDisc1",
                     "sheet": "Escapes",
                     "column": 33,
                     "substract": False,
                     "summary_column_list": columnasEscapes,
                     "offset": 6}},
    # Respuestas nosepoke forzados no discriminativos 2
    {"conteototal": {"mark1": 308,
                     "label": "EscLibNoDisc2",
                     "sheet": "Escapes",
                     "column": 35,
                     "substract": False,
                     "summary_column_list": columnasEscapes,
                     "offset": 7}},

    # Latencias Escape
    # Latencias nosepoke forzados discriminativos reforzados
    {"conteolat": {"mark1": 114, "mark2": 301,
                   "label": "LatEscFDiscRef",
                   "sheet": "LatNosepoke",
                   "column": 37,
                   "substract": False,
                   "summary_column_list": columnasEscapes,
                   "offset": 0}},
    # Latencias nosepoke forzados discriminativos no reforzados
    {"conteolat": {"mark1": 115, "mark2": 302,
                   "label": "LatEscFDiscNoRef",
                   "sheet": "LatNosepoke",
                   "column": 39,
                   "substract": False,
                   "summary_column_list": columnasEscapes,
                   "offset": 1}},
    # Latencias nosepoke forzados no discriminativos 1
    {"conteolat": {"mark1": 134, "mark2": 303,
                   "label": "LatEscFNoDisc1",
                   "sheet": "LatNosepoke",
                   "column": 41,
                   "substract": False,
                   "summary_column_list": columnasEscapes,
                   "offset": 2}},
    # Latencias nosepoke forzados no discriminativos 2
    {"conteolat": {"mark1": 137, "mark2": 304,
                   "label": "LatEscFNoDisc2",
                   "sheet": "LatNosepoke",
                   "column": 43,
                   "substract": False,
                   "summary_column_list": columnasEscapes,
                   "offset": 3}},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"mark1": 154, "mark2": 305,
                   "label": "LatEscLDiscRef",
                   "sheet": "LatNosepoke",
                   "column": 45,
                   "substract": False,
                   "summary_column_list": columnasEscapes,
                   "offset": 4}},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"mark1": 155, "mark2": 306,
                   "label": "LatEscLDiscRef",
                   "sheet": "LatNosepoke",
                   "column": 47,
                   "substract": False,
                   "summary_column_list": columnasEscapes,
                   "offset": 5}},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"mark1": 157, "mark2": 307,
                   "label": "LatEscLNoDisc1",
                   "sheet": "LatNosepoke",
                   "column": 49,
                   "substract": False,
                   "summary_column_list": columnasEscapes,
                   "offset": 6}},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"mark1": 160, "mark2": 308,
                   "label": "LatEscLDisc2",
                   "sheet": "LatNosepoke",
                   "column": 51,
                   "substract": False,
                   "summary_column_list": columnasEscapes,
                   "offset": 7}},

    # Escapes forzados
    # Respuestas nosepoke escape forzado discriminativo
    {"conteototal": {"mark1": 403,
                     "label": "EscForzDisc",
                     "sheet": "EscapesForzados",
                     "column": 53,
                     "substract": False,
                     "summary_column_list": columnasEscForz,
                     "offset": 0}},
    # Respuestas nosepoke escape forzado no discriminativo
    {"conteototal": {"mark1": 406,
                     "label": "EscForzNoDisc",
                     "sheet": "EscapesForzados",
                     "column": 55,
                     "substract": False,
                     "summary_column_list": columnasEscForz,
                     "offset": 1}},

    # Latencias Escape Forzado
    # Latencias nosepoke escape forzado discriminativo
    {"conteolat": {"mark1": 401, "mark2": 403,
                   "label": "LatEscForzDisc",
                   "sheet": "LatEscapeForz",
                   "column": 57,
                   "substract": False,
                   "summary_column_list": columnasLatPal,
                   "offset": 0}},
    # Latencias nosepoke escape forzado no discriminativo
    {"conteolat": {"mark1": 404, "mark2": 406,
                   "label": "LatEscForzNoDisc",
                   "sheet": "LatEscapeForz",
                   "column": 59,
                   "substract": False,
                   "summary_column_list": columnasLatPal,
                   "offset": 1}},

    # Latencias Escape Forzado por estímulo
    # Latencias nosepoke escape forzado discriminativo positivo
    {"conteolat": {"mark1": 407, "mark2": 403,
                   "label": "LatEscForzDiscPos",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 61,
                   "substract": False,
                   "summary_column_list": columnasLatPal,
                   "offset": 0}},
    # Latencias nosepoke escape forzado no discriminativo negativo
    {"conteolat": {"mark1": 408, "mark2": 403,
                   "label": "LatEscForzDiscNeg",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 63,
                   "substract": False,
                   "summary_column_list": columnasLatPal,
                   "offset": 1}},
    # Latencias nosepoke escape forzado discriminativo 1
    {"conteolat": {"mark1": 409, "mark2": 406,
                   "label": "LatEscForzNoDiscLuz1",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 65,
                   "substract": False,
                   "summary_column_list": columnasLatPal,
                   "offset": 2}},
    # Latencias nosepoke escape forzado no discriminativo 2
    {"conteolat": {"mark1": 410, "mark2": 406,
                   "label": "LatEscForzNoDiscLuz2",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 67,
                   "substract": False,
                   "summary_column_list": columnasLatPal,
                   "offset": 3}},

    # Análisis múltiples
    # Latencias Pal Disc
    {"agregate": {"conteolat": {"measures": 2,
                                "mark1": 112, "mark2": 113,
                                "mark3": 401, "mark4": 402,
                                "label": "LatPalDiscTotal",
                                "sheet": "Latencias",
                                "column": 69,
                                "substract": False,
                                "summary_column_list": columnasLatPal,
                                "offset": 0}}},
    # Latencias Pal No Disc
    {"agregate": {"conteolat": {"measures": 2,
                                "mark1": 132, "mark2": 133,
                                "mark3": 404, "mark4": 405,
                                "label": "LatPalNoDiscTotal",
                                "sheet": "Latencias",
                                "column": 71,
                                "substract": False,
                                "summary_column_list": columnasLatPal,
                                "offset": 1}}},
]

purgeSessions(directorioTemporal, sujetos, sesionesPresentes, columnasProp, columnasResp, columnasLatPal,
              columnasEscapes, columnasLatEsc, columnasEscForz)

with open("data.json", "w") as data_file:
    json.dump(analysis_list, data_file, indent=4)
with open("data.json", "r") as data_file:
    json_data = json.load(data_file)

convertir(directorioTemporal, directorioBrutos, directorioConvertidos, sujetos, sesionesPresentes, subfijo="_SUBCHOIL_",
          mover=False)

wb = createDocument(archivo, directorioConvertidos)

sheet_dict = create_sheets(wb, 'Proporciones', 'Respuestas', 'Latencias', 'Comedero', 'Escapes', 'LatNosepoke',
                           'EscapesForzados', 'LatEscapeForz', 'LatEscapeForzPorEstim')

analyze(dirConv=directorioConvertidos, fileName=archivo, subList=sujetos, sessionList=sesionesPresentes,
        suffix="_SUBCHOIL_", workbook=wb, sheetDict=sheet_dict, analysisList=analysis_list, markColumn="P",
        timeColumn="O")
