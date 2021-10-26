import json
from Funciones import *

archivo = 'Prueba2.xlsx'
# directorioBrutos = 'C:/Users/Admin/Desktop/Escape/Alter/2021_Alter_S/Brutos/'
# directorioConvertidos = 'C:/Users/Admin/Desktop/Escape/Alter/2021_Alter_S/Convertidos/'
# directorioTemporal = 'C:/Users/Admin/Desktop/Escape/Alter/2021_Alter_S/Temporal/'
directorioTemporal = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Temporal/'
directorioBrutos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/PruebasBrutos/'
directorioConvertidos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Flex/'

sujetos = ['S3', 'S4', 'S5', 'S6', 'S7', 'S8', 'S9', 'S10']
columnas_latencias = {"S3": 2, "S4": 7, "S5": 12, "S6": 17, "S7": 22, "S8": 27, "S9": 32, "S10": 37}
sesionesPresentes = []  # Lista vacía.
marcadores = []  # Lista vacía.
tiempo = []  # Lista vacía.

analysis_list = [
    # Ensayos completados
    {"fetch": {"sheet": "Latencias",
               "summary_column_list": columnas_latencias,
               "cell_row": 12,
               "cell_column": 2,
               "offset": 0,
               }},

    # Análisis múltiples
    # Latencias Pal Izq
    {"conteolat": {"measures": 3,
                   "mark1": 202, "mark2": 303,  # Palanca Izquierda Blanca
                   "mark3": 202, "mark4": 306,  # Palanca Izquierda Roja
                   "mark5": 202, "mark6": 309,  # Palanca Izquierda Azul
                   "header": "LatPalIzqTotal",
                   "column": 1,
                   "sheet": "Latencias",
                   "summary_column_list": columnas_latencias,
                   "offset": 1}},

    # Latencias Pal Der
    {"conteolat": {"measures": 3,
                   "mark1": 206, "mark2": 312,  # Palanca Derecha Blanca
                   "mark3": 206, "mark4": 315,  # Palanca Derecha Roja
                   "mark5": 206, "mark6": 318,  # Palanca Derecha Azul
                   "header": "LatPalDerTotal",
                   "column": 3,
                   "sheet": "Latencias",
                   "summary_column_list": columnas_latencias,
                   "offset": 2}},

    # Latencias Nosepoke
    {"conteolat": {"mark1": 210, "mark2": 321,
                   "header": "LatNosepoke",
                   "column": 5,
                   "sheet": "Latencias",
                   "summary_column_list": columnas_latencias,
                   "offset": 3}},
]

with open("alter_funciones_data.json", "r") as data_file:
    json_data = json.load(data_file)

purgeSessions(directorioTemporal, sujetos, sesionesPresentes)

convertir(dirTemp=directorioTemporal, dirPerm=directorioBrutos, dirConv=directorioConvertidos, subjectList=sujetos,
          presentSessions=sesionesPresentes, subfijo="_ALT_", mover=False)

wb = createDocument(fileName=archivo, targetDirectory=directorioConvertidos)

sheet_dict = create_sheets(wb, "Latencias")

analyze(dirConv=directorioConvertidos, fileName=archivo, subList=sujetos, sessionList=sesionesPresentes,
        suffix="_ALT_", workbook=wb, sheetDict=sheet_dict, analysisList=json_data, markColumn="L",
        timeColumn="K")
