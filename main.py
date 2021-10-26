from Funciones import *

archivo = 'Prueba_enumerate.xlsx'
# directorioBrutos = 'C:/Users/Admin/Desktop/Escape/Datos/Brutos/'
# directorioConvertidos = 'C:/Users/Admin/Desktop/Escape/Datos/ConvertidosPython/Escape/'
directorioTemporal = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Temporal/'
directorioBrutos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/PruebasBrutos/'
directorioConvertidos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Flex/'

sujetos = ['E3', 'E4', 'E5', 'E7', 'E8', 'E9', 'E1']
columnasProp = {"E3": 2, "E4": 4, "E5": 6, "E7": 8, "E8": 10, "E9": 12, "E1": 14}
columnasResp = {"E3": 2, "E4": 9, "E5": 16, "E7": 23, "E8": 30, "E9": 37, "E1": 44}
columnasLatPal = {"E3": 2, "E4": 7, "E5": 12, "E7": 17, "E8": 22, "E9": 27, "E1": 32}
columnasEscapes = {"E3": 2, "E4": 13, "E5": 24, "E7": 35, "E8": 46, "E9": 57, "E1": 68}
columnasLatEsc = {"E3": 2, "E4": 13, "E5": 24, "E7": 35, "E8": 46, "E9": 57, "E1": 68}
columnasEscForz = {"E3": 2, "E4": 7, "E5": 12, "E7": 17, "E8": 22, "E9": 27, "E1": 32}

# sujetos = ["S3", "S4", "S5", "S6", "S7", "S8", "S9", "S10"]
# columnasProp = {"S3": 2, "S4": 4, "S5": 6, "S6": 8, "S7": 10, "S8": 12, "S9": 14, "S10": 16}
# columnasResp = {"S3": 2, "S4": 9, "S5": 16, "S6": 23, "S7": 30, "S8": 37, "S9": 44, "S10": 51}
# columnasLatPal = {"S3": 2, "S4": 7, "S5": 12, "S6": 17, "S7": 22, "S8": 27, "S9": 32, "S10": 37}
# columnasEscapes = {"S3": 2, "S4": 13, "S5": 24, "S6": 35, "S7": 46, "S8": 57, "S9": 68, "S10": 79}
# columnasLatEsc = {"S3": 2, "S4": 13, "S5": 24, "S6": 35, "S7": 46, "S8": 57, "S9": 68, "S10": 79}
# columnasEscForz = {"S3": 2, "S4": 7, "S5": 12, "S6": 17, "S7": 22, "S8": 27, "S9": 32, "S10": 37}
sesionesPresentes = []  # Lista vacía.
marcadores = []  # Lista vacía.
tiempo = []  # Lista vacía.

analysis_list = [
    # Proporciones
    {"fetch": {"cell_row": 14,
               "cell_column": 6,
               "sheet": "Proporciones",
               "summary_column_list": columnasProp,
               "offset": 0
               }},
    # Ensayos completados
    {"fetch": {"cell_row": 15,
               "cell_column": 2,
               "sheet": "Proporciones",
               "summary_column_list": columnasProp,
               "offset": 1
               }},

    # Respuestas palancas
    # Respuestas Pal forzados discriminativos reforzados
    {"conteoresp": {"mark1": 114, "mark2": 180, "mark3": 202,
                    "header": "PalDiscRef",
                    "sheet": "Respuestas",
                    "column": 1,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 0}},
    # Respuestas Pal forzados discriminativos no reforzados
    {"conteoresp": {"mark1": 115, "mark2": 180, "mark3": 202,
                    "header": "PalNoDiscRef",
                    "sheet": "Respuestas",
                    "column": 3,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 1}},
    # Respuestas Pal forzados no discriminativos 1
    {"conteoresp": {"mark1": 134, "mark2": 180, "mark3": 201,
                    "header": "PalNoDisc1",
                    "sheet": "Respuestas",
                    "column": 5,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 2}},
    # Respuestas Pal forzados no discriminativos 2
    {"conteoresp": {"mark1": 137, "mark2": 180, "mark3": 201,
                    "header": "PalNoDisc2",
                    "sheet": "Respuestas",
                    "column": 7,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 3}},

    # Latencias palancas
    # Latencias pal forzados discriminativos
    {"conteolat": {"mark1": 112, "mark2": 113,
                   "header": "LatPalDisc",
                   "sheet": "Latencias",
                   "column": 9,
                   "summary_column_list": columnasLatPal,
                   "offset": 0}},
    # Latencias pal forzados no discriminativos
    {"conteolat": {"mark1": 132, "mark2": 133,
                   "header": "LatPalNoDisc",
                   "sheet": "Latencias",
                   "column": 11,
                   "summary_column_list": columnasLatPal,
                   "offset": 1}},

    # Respuestas comederos
    # Respuestas comederos forzados discriminativos reforzados
    {"conteoresp": {"mark1": 114, "mark2": 16, "mark3": 203,
                    "header": "ComDiscRef",
                    "sheet": "Comedero",
                    "column": 13,
                    "summary_column_list": columnasResp,
                    "offset": 0}},
    # Respuestas comederos forzados discriminativos no reforzados
    {"conteoresp": {"mark1": 115, "mark2": 117, "mark3": 203,
                    "header": "ComNoDiscRef",
                    "sheet": "Comedero",
                    "column": 15,
                    "summary_column_list": columnasResp,
                    "offset": 1}},
    # Respuestas comederos forzados no discriminativos 1
    {"conteoresp": {"mark1": 134, "mark2": 40, "mark3": 203,
                    "header": "ComNoDisc1",
                    "sheet": "Comedero",
                    "column": 17,
                    "summary_column_list": columnasResp,
                    "offset": 2}},
    # Respuestas comederos forzados no discriminativos 2
    {"conteoresp": {"mark1": 137, "mark2": 43, "mark3": 203,
                    "header": "ComNoDisc2",
                    "sheet": "Comedero",
                    "column": 19,
                    "summary_column_list": columnasResp,
                    "offset": 3}},

    # Escapes
    # Respuestas nosepoke forzados discriminativos reforzados
    {"conteototal": {"mark1": 301,
                     "header": "EscForzDiscRef",
                     "sheet": "Escapes",
                     "column": 21,
                     "summary_column_list": columnasEscapes,
                     "offset": 0}},
    # Respuestas nosepoke forzados discriminativos no reforzados
    {"conteototal": {"mark1": 302,
                     "header": "EscForzDiscNoRef",
                     "sheet": "Escapes",
                     "column": 23,
                     "summary_column_list": columnasEscapes,
                     "offset": 1}},
    # Respuestas nosepoke forzados no discriminativos 1
    {"conteototal": {"mark1": 303,
                     "header": "EscForzNoDisc1",
                     "sheet": "Escapes",
                     "column": 25,
                     "summary_column_list": columnasEscapes,
                     "offset": 2}},
    # Respuestas nosepoke forzados no discriminativos 2
    {"conteototal": {"mark1": 304,
                     "header": "EscForzNoDisc2",
                     "sheet": "Escapes",
                     "column": 27,
                     "summary_column_list": columnasEscapes,
                     "offset": 3}},
    # Respuestas nosepoke forzados discriminativos reforzados
    {"conteototal": {"mark1": 305,
                     "header": "EscLibDiscRef",
                     "sheet": "Escapes",
                     "column": 29,
                     "summary_column_list": columnasEscapes,
                     "offset": 4}},
    # Respuestas nosepoke forzados discriminativos no reforzados
    {"conteototal": {"mark1": 306,
                     "header": "EscLibDiscNoRef",
                     "sheet": "Escapes",
                     "column": 31,
                     "summary_column_list": columnasEscapes,
                     "offset": 5}},
    # Respuestas nosepoke forzados no discriminativos 1
    {"conteototal": {"mark1": 307,
                     "header": "EscLibNoDisc1",
                     "sheet": "Escapes",
                     "column": 33,
                     "summary_column_list": columnasEscapes,
                     "offset": 6}},
    # Respuestas nosepoke forzados no discriminativos 2
    {"conteototal": {"mark1": 308,
                     "header": "EscLibNoDisc2",
                     "sheet": "Escapes",
                     "column": 35,
                     "summary_column_list": columnasEscapes,
                     "offset": 7}},

    # Latencias Escape
    # Latencias nosepoke forzados discriminativos reforzados
    {"conteolat": {"mark1": 114, "mark2": 301,
                   "header": "LatEscFDiscRef",
                   "sheet": "LatNosepoke",
                   "column": 37,
                   "summary_column_list": columnasEscapes,
                   "offset": 0}},
    # Latencias nosepoke forzados discriminativos no reforzados
    {"conteolat": {"mark1": 115, "mark2": 302,
                   "header": "LatEscFDiscNoRef",
                   "sheet": "LatNosepoke",
                   "column": 39,
                   "summary_column_list": columnasEscapes,
                   "offset": 1}},
    # Latencias nosepoke forzados no discriminativos 1
    {"conteolat": {"mark1": 134, "mark2": 303,
                   "header": "LatEscFNoDisc1",
                   "sheet": "LatNosepoke",
                   "column": 41,
                   "summary_column_list": columnasEscapes,
                   "offset": 2}},
    # Latencias nosepoke forzados no discriminativos 2
    {"conteolat": {"mark1": 137, "mark2": 304,
                   "header": "LatEscFNoDisc2",
                   "sheet": "LatNosepoke",
                   "column": 43,
                   "summary_column_list": columnasEscapes,
                   "offset": 3}},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"mark1": 154, "mark2": 305,
                   "header": "LatEscLDiscRef",
                   "sheet": "LatNosepoke",
                   "column": 45,
                   "summary_column_list": columnasEscapes,
                   "offset": 4}},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"mark1": 155, "mark2": 306,
                   "header": "LatEscLDiscRef",
                   "sheet": "LatNosepoke",
                   "column": 47,
                   "summary_column_list": columnasEscapes,
                   "offset": 5}},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"mark1": 157, "mark2": 307,
                   "header": "LatEscLNoDisc1",
                   "sheet": "LatNosepoke",
                   "column": 49,
                   "summary_column_list": columnasEscapes,
                   "offset": 6}},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"mark1": 160, "mark2": 308,
                   "header": "LatEscLDisc2",
                   "sheet": "LatNosepoke",
                   "column": 51,
                   "summary_column_list": columnasEscapes,
                   "offset": 7}},

    # Escapes forzados
    # Respuestas nosepoke escape forzado discriminativo
    {"conteototal": {"mark1": 403,
                     "header": "EscForzDisc",
                     "sheet": "EscapesForzados",
                     "column": 53,
                     "summary_column_list": columnasEscForz,
                     "offset": 0}},
    # Respuestas nosepoke escape forzado no discriminativo
    {"conteototal": {"mark1": 406,
                     "header": "EscForzNoDisc",
                     "sheet": "EscapesForzados",
                     "column": 55,
                     "summary_column_list": columnasEscForz,
                     "offset": 1}},

    # Latencias Escape Forzado
    # Latencias nosepoke escape forzado discriminativo
    {"conteolat": {"mark1": 401, "mark2": 403,
                   "header": "LatEscForzDisc",
                   "sheet": "LatEscapeForz",
                   "column": 57,
                   "summary_column_list": columnasLatPal,
                   "offset": 0}},
    # Latencias nosepoke escape forzado no discriminativo
    {"conteolat": {"mark1": 404, "mark2": 406,
                   "header": "LatEscForzNoDisc",
                   "sheet": "LatEscapeForz",
                   "column": 59,
                   "summary_column_list": columnasLatPal,
                   "offset": 1}},

    # Latencias Escape Forzado por estímulo
    # Latencias nosepoke escape forzado discriminativo positivo
    {"conteolat": {"mark1": 407, "mark2": 403,
                   "header": "LatEscForzDiscPos",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 61,
                   "summary_column_list": columnasLatPal,
                   "offset": 0}},
    # Latencias nosepoke escape forzado no discriminativo negativo
    {"conteolat": {"mark1": 408, "mark2": 403,
                   "header": "LatEscForzDiscNeg",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 63,
                   "summary_column_list": columnasLatPal,
                   "offset": 1}},
    # Latencias nosepoke escape forzado discriminativo 1
    {"conteolat": {"mark1": 409, "mark2": 406,
                   "header": "LatEscForzNoDiscLuz1",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 65,
                   "summary_column_list": columnasLatPal,
                   "offset": 2}},
    # Latencias nosepoke escape forzado no discriminativo 2
    {"conteolat": {"mark1": 410, "mark2": 406,
                   "header": "LatEscForzNoDiscLuz2",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 67,
                   "summary_column_list": columnasLatPal,
                   "offset": 3}},

    # Análisis múltiples
    # Latencias Pal Disc
    {"conteolat": {"measures": 2,
                   "mark1": 112, "mark2": 113,  # Ensayos forzados discriminativos
                   "mark3": 401, "mark4": 402,  # Ensayos discriminativos de escape forzado
                   "header": "LatPalDiscTotal",
                   "sheet": "Latencias",
                   "column": 69,
                   "summary_column_list": columnasLatPal,
                   "offset": 0}},
    # Latencias Pal No Disc
    {"conteolat": {"measures": 2,
                   "mark1": 132, "mark2": 133,  # Ensayos forzados no discriminativos
                   "mark3": 404, "mark4": 405,  # Ensayos no discriminativos de escape forzado
                   "header": "LatPalNoDiscTotal",
                   "sheet": "Latencias",
                   "column": 71,
                   "summary_column_list": columnasLatPal,
                   "offset": 1}},
]

purgeSessions(directorioTemporal, sujetos, sesionesPresentes)

convertir(dirTemp=directorioTemporal, dirPerm=directorioBrutos, dirConv=directorioConvertidos, subjectList=sujetos,
          presentSessions=sesionesPresentes, subfijo="_SUBCHOIL_", mover=False)

wb = createDocument(fileName=archivo, targetDirectory=directorioConvertidos)

sheet_dict = create_sheets(wb, 'Proporciones', 'Respuestas', 'Latencias', 'Comedero', 'Escapes', 'LatNosepoke',
                           'EscapesForzados', 'LatEscapeForz', 'LatEscapeForzPorEstim')

analyze(dirConv=directorioConvertidos, fileName=archivo, subList=sujetos, sessionList=sesionesPresentes,
        suffix="_SUBCHOIL_", workbook=wb, sheetDict=sheet_dict, analysisList=analysis_list, markColumn="P",
        timeColumn="O")
