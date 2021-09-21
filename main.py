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
sesionesPresentes = []  # Esta lista debe estar vacía.
marcadores = []
tiempo = []
analysis_list = [{"conteoresp": (114, 180, 202, "resPalForzDiscRef", 1, True)},
                 {"conteorespa": (444, 555, 666)},
                 {"conteolat": (123, 234)}
                 ]

purgeSessions(directorioTemporal, sujetos, sesionesPresentes, columnasProp, columnasResp, columnasLatPal,
              columnasEscapes, columnasLatEsc, columnasEscForz)
print("Purged")

convertir(directorioTemporal, directorioBrutos, directorioConvertidos, sujetos, sesionesPresentes, subfijo="_SUBCHOIL_")
print("Converted")

wb = createDocument(archivo, directorioConvertidos)

sheet_list = create_sheets(wb, 'Proporciones', 'Respuestas', 'Latencias', 'Comedero', 'Escapes', 'LatNosepoke',
                           'EscapesForzados', 'LatEscapeForz')

analyze(dirConv=directorioConvertidos, fileName=archivo, subList=sujetos, sessionList=sesionesPresentes,
        suffix="_SUBCHOIL_", workbook=wb, sheetList=sheet_list, analysisList=analysis_list, markColumn="P",
        timeColumn="O")

wb.save(directorioConvertidos + archivo)
