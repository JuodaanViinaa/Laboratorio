from Funciones import *

archivo = 'Response_distribution.xlsx'
# directorioBrutos = 'C:/Users/Admin/Desktop/Escape/Datos/Brutos/'
# directorioConvertidos = 'C:/Users/Admin/Desktop/Escape/Datos/ConvertidosPython/Escape/'
directorioTemporal = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Distribucion_respuestas/'
directorioBrutos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/'
directorioConvertidos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Distribucion_respuestas/convertidos/'

sujetos = ["A1", "A2", "A3"]
columnas = {"A1": 2}

sesionesPresentes = []  # Lista vacía.
marcadores = []  # Lista vacía.
tiempo = []  # Lista vacía.

analysis_list = [
    {"resp_dist": {"mark1": 300, "mark2": 300, "mark3": 200,
                   "bin_size": 1,
                   "bin_amount": 15,
                   }},
]

purgeSessions(directorioTemporal, sujetos, sesionesPresentes)

convertir(dirTemp=directorioTemporal, dirPerm=directorioBrutos, dirConv=directorioConvertidos, subjectList=sujetos,
          presentSessions=sesionesPresentes, mover=False)

wb = createDocument(fileName=archivo, targetDirectory=directorioConvertidos)

sheet_dict = create_sheets(wb, "Placeholder")

analyze(dirConv=directorioConvertidos, fileName=archivo, subList=sujetos, sessionList=sesionesPresentes,
        workbook=wb, sheetDict=sheet_dict, suffix="_", analysisList=analysis_list, markColumn="N", timeColumn="M")
