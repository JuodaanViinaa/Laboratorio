from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from os import listdir
from statistics import mean, median
import pandas
import re
from shutil import move


# archivo = 'ResumenForz.xlsx'
# # directorioBrutos = 'C:/Users/Admin/Desktop/Escape/Datos/Brutos/'
# # directorioConvertidos = 'C:/Users/Admin/Desktop/Escape/Datos/ConvertidosPython/Escape/'
# directorioTemporal = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Temporal/'
# directorioBrutos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/PruebasBrutos/'
# directorioConvertidos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Flex/'
#
# sujetos = ['E3', 'E4', 'E5', 'E7', 'E8', 'E9', 'E1']
# columnasProp = [2, 3, 4, 5, 6, 7, 8]
# columnasResp = [2, 9, 16, 23, 30, 37, 44]
# columnasLatPal = [2, 7, 12, 17, 22, 27, 32]
# columnasEscapes = [2, 13, 24, 35, 46, 57, 68]
# columnasLatEsc = [2, 13, 24, 35, 46, 57, 68]
# columnasEscForz = [2, 7, 12, 17, 22, 27, 32]
# sesionesPresentes = []  # Esta lista debe estar vacía.


# Función para determinar el número de sesiones. Se leen los archivos del directorio temporal y con base en sus
# nombres se determinan los sujetos presentes y sus sesiones.
def purgeSessions(temporaryDirectory, subjectList, sessionList, *columnLists):
    """
    Se genera una lista temporal que contiene aquellos sujetos cuyos datos sí están en la carpeta temporal; después, se
    eliminan las columnas pertinentes de cada una de las listas de columnas. Los argumentos son: el directorio temporal
    en que están los datos brutos, una lista con los nombres de los sujetos, y las listas de columnas.
    :param temporaryDirectory: Directorio donde se almacenan temporalmente los datos brutos por analizar.
    :param subjectList: Lista con los nombres de todos los sujetos a analizar.
    :param sessionList: Lista vacía que contendrá las sesiones presentes por analizar para cada sujeto.
    :param columnLists: Listas con los valores de las columnas en que se pegarán los datos para analizar.
    :return:
    """
    sujetosFaltantes = []
    dirTemp = sorted(listdir(temporaryDirectory))
    listaTemp = []
    for sbj in subjectList:
        for datoTemp in dirTemp:
            if sbj == datoTemp.split('_')[0] and sbj not in listaTemp:
                listaTemp.append(sbj)

    # Los sujetos que no forman parte de la lista temporal son agregados a la lista sujetosFaltantes para que sus
    # columnas sean eliminadas también del análisis.
    for sbj in subjectList:
        if sbj not in listaTemp:
            sujetosFaltantes.append(sbj)
    # Si faltan sujetos se imprime quiénes son. Si no, se indica con un mensaje.
    if len(sujetosFaltantes) == 0:
        print("Todos los sujetos tienen al menos una sesión por analizar.")
    else:
        print(f"Sujetos faltantes: {str(sujetosFaltantes)}")

    # Se hace una lista de listas con las sesiones presentes de cada sujeto.
    # El código compara el nombre de cada uno de los sujetos que sí tienen sesiones con el nombre de cada uno de los
    # archivos del directorio temporal. Si encuentra una coincidencia, toma el número de la sesión asociado con el
    # archivo encontrado y lo agrega a una sublista que contendrá todas las sesiones del sujeto en cuestión. El código
    # tolera sesiones salteadas
    indice = 0
    for sujetoPresente in listaTemp:
        sessionList.append([])
        sublista = []
        for datoTemp in dirTemp:
            if sujetoPresente == datoTemp.split('_')[0]:
                sublista.append(int(datoTemp.split('_')[-1]))
        sessionList[indice] = sorted(sublista)
        indice += 1

    for i in range(len(listaTemp)):
        print(f"Sesiones presentes del sujeto {listaTemp[i]}: {sessionList[i]}\n")

    # Se eliminan los elementos pertinentes de las listas de columnas si algún sujeto falta.
    for sujetoFaltante in sujetosFaltantes:
        for columnList in columnLists:
            if sujetoFaltante in subjectList:
                del columnList[subjectList.index(sujetoFaltante)]
        del subjectList[subjectList.index(sujetoFaltante)]


# Convertidor
def convertir(dirTemp, dirPerm, dirConv, subjectList, presentSessions, columnas=6, subfijo='_'):
    """
    Convierte archivos de texto plano de MedPC en hojas de cálculo en formato *.xlsx.
    Separa cada lista en dos columnas con base en el punto decimal.
    :param dirTemp: Directorio donde se almacenan temporalmente los datos brutos por analizar.
    :param dirPerm: Directorio donde se almacenarán finalmente los datos brutos después de la conversión.
    :param dirConv: Directorio donde se almacenarán los archivos ya convertidos.
    :param subjectList: Lista con los nombres de todos los sujetos.
    :param presentSessions: Lista vacía rellenada por el programa con las sesiones presentes de cada sujeto.
    :param columnas: Cantidad de columnas en que están divididos los archivos de texto de Med. Valor por defecto: 6.
    :param subfijo: Identificador del nombre de los archivos. Ejemplo: '_Alter_'. Valor por defecto: ''.
    :return:
    """
    for sjt in range(len(subjectList)):
        print(len(subjectList))
        print(presentSessions)
        for ssn in presentSessions[sjt]:
            print(f"Convirtiendo sesión {str(ssn)} de sujeto {subjectList[sjt]}.")
            # Pandas lee los datos y los escribe en el archivo convertido en 6 columnas separando por los espacios.
            # El argumento names indica cuántas columnas se crearán. Evita errores cuando se edita el archivo de Med.
            datos = pandas.read_csv(dirTemp + subjectList[sjt] + subfijo + str(ssn), header=None,
                                    names=range(columnas), sep=r'\s+')  # ¿Por qué utilizo range(columnas)?
            datos.to_excel(dirConv + subjectList[sjt] + subfijo + str(ssn) + '.xlsx', index=False,
                           header=None)

            # Openpyxl abre el archivo creado por pandas, lee la hoja y la almacena en la variable hojaCompleta.
            archivoCompleto = load_workbook(dirConv + subjectList[sjt] + subfijo + str(ssn) + '.xlsx')
            hojaCompleta = archivoCompleto.active

            # Se genera una lista que contenga sub-listas con todos los valores de las listas dadas por Med.
            # Funciona para cualquier cantidad de listas.
            # Los datos se convierten en flotantes para que todos tengan punto decimal, y luego en string para que mas
            # adelante el método split los pueda separar por el punto.
            metalista = [[]]
            contadormetalista = 0

            columna1 = hojaCompleta['B']
            for fila in range(12, len(columna1)):
                for columna in range(2, columnas + 1):
                    if hojaCompleta[get_column_letter(columna) + str(fila)].value is not None:
                        metalista[contadormetalista].append(str(float(hojaCompleta[get_column_letter(columna) +
                                                                                   str(fila)].value)))
                    elif hojaCompleta[get_column_letter(columna) + str(fila)].value is None and columna == 2:
                        metalista.append([])
                        contadormetalista += 1

            # Escribir cada sub-lista en una columna de excel separando por punto decimal.
            # Se utilizan expresiones regulares (regex) para indicar al programa que debe añadir ceros cuando pandas
            # los ha eliminado (cuando están al final de una cifra después de un punto decimal).
            regex1 = re.compile(r'^\d+\.\d{2}$')
            regex2 = re.compile(r'^\d+\.\d$')
            for ii in range(len(metalista)):
                for j in range(len(metalista[ii])):
                    if regex1.search(metalista[ii][j]):
                        metalista[ii][j] += '0'
                    elif regex2.search(metalista[ii][j]):  # Esto está por probarse. No sé si necesito agregar '00'.
                        metalista[ii][j] += '00'
                    hojaCompleta[get_column_letter((ii * 2) + columnas + 3) + str(j + 1)] = int(
                        metalista[ii][j].split('.')[0])
                    hojaCompleta[get_column_letter((ii * 2) + columnas + 4) + str(j + 1)] = int(
                        metalista[ii][j].split('.')[1])
            archivoCompleto.save(dirConv + subjectList[sjt] + subfijo + str(ssn) + '.xlsx')
            move(dirTemp + subjectList[sjt] + subfijo + str(ssn), dirPerm + subjectList[sjt] + subfijo + str(ssn))
        print('\n')


def createDocument(fileName, targetDirectory):
    # Revisar si el archivo de resumen ya existe. De lo contrario, crearlo.
    if fileName in listdir(targetDirectory):
        print('Summary file found. Opening...')
        wb = load_workbook(targetDirectory + fileName)
    else:
        print('Summary file not found. Creating...')
        wb = Workbook()
    return wb


def create_sheets(workbook, *sheets):
    sheet_list = []
    for sheet in sheets:
        if sheet not in workbook.sheetnames:
            new_sheet = workbook.create_sheet(sheet)
        else:
            new_sheet = workbook[sheet]
        sheet_list.append(new_sheet)
    return sheet_list
    # Trabajar más adelante en el código con una lista de hojas. Accesarlas por su índice con una cosa bien redundante:
    # sheet_list[sheet_list.index('respuestas')]
    # O usar un diccionario y acceder por key


# Función para contar respuestas por tipo de ensayo. Los argumentos son marcadores de Med.
def conteoresp(marcadores, inicioEnsayo, finEnsayo, respuesta):
    contadorTemp = 0
    inicio = 0
    resp = []
    for n in range(1, len(marcadores)):
        if marcadores[n].value == inicioEnsayo:
            inicio = 1
        elif marcadores[n].value == respuesta and inicio == 1:
            contadorTemp += 1
        elif marcadores[n].value == finEnsayo and inicio == 1:
            inicio = 0
            resp.append(contadorTemp)
            contadorTemp = 0
    if len(resp) == 0:
        resp = [0]
    return resp


# Función para contar respuestas totales. El argumento es el marcador de Med.
def conteototal(marcadores, respuesta):
    contador = 0
    for n in range(len(marcadores)):
        if marcadores[n].value == respuesta:
            contador += 1
    return contador


# Función para contar latencias. Si en un ensayo no hay respuestas que contar, la función resulta en una lista con un
# cero. Los argumentos son marcadores de Med.
def conteolat(marcadores, tiempo, inicioensayo, respuesta):
    inicio = 0
    lat = []
    tiempoini = 0
    for n in range(1, len(marcadores)):
        if marcadores[n].value == inicioensayo:
            inicio = 1
            tiempoini = tiempo[n].value
        elif marcadores[n].value == respuesta and inicio == 1:
            lat.append((tiempo[n].value - tiempoini) / 20)
            inicio = 0
    if len(lat) == 0:
        lat = [0]
    return lat


# Función para escribir listas en columnas. El argumento "restar" indica si se debe restar uno a las respuestas dadas
# en cada ensayo. Esto solo sucede en las palancas dado que se registra también la respuesta que le da inicio al ensayo.
def esccolumnas(hojaind, titulo, columna, lista, restar):
    hojaind[get_column_letter(columna) + str(1)] = titulo
    if restar:
        for pos in range(len(lista)):
            if lista[pos] > 0:
                hojaind[get_column_letter(columna) + str(pos + 2)] = lista[pos] - 1
            else:
                hojaind[get_column_letter(columna) + str(pos + 2)] = lista[pos]
    else:
        for pos in range(len(lista)):
            hojaind[get_column_letter(columna) + str(pos + 2)] = lista[pos]


analysis_list = [{"conteoresp": (114, 180, 202, "resPalForzDiscRef", 0, 1, True)},
                 # hoja: (marcadores, "etiquetaParaIndividual", posicionHojaEnLista, columnaParaIndividual, ¿Restar?
                 {"conteoresp": (444, 555, 666)},
                 {"conteolat": (123, 234)}
                 ]


def analyze(dirConv, fileName, subList, sessionList, suffix, workbook, sheetList, analysisList, markColumn, timeColumn):
    for subject in range(len(subList)):
        print(f"Trying subject {subList[subject]}.")
        for session in sessionList[subject]:
            print(f"Trying session {session}.")
            sujetoWb = load_workbook(dirConv + subList[subject] + suffix + str(session) + '.xlsx')
            sujetoWs = sujetoWb.worksheets[0]
            hojaind = sujetoWb.create_sheet('FullLists')
            tiempo = sujetoWs[timeColumn]
            marcadores = sujetoWs[markColumn]

            for analysis in analysisList:
                key, value = list(analysis.items())[0]
                if key == "conteoresp":
                    print(value)
                    respuestas = conteoresp(marcadores, value[0], value[1], value[2])
                    sheetList[value[4]][
                        get_column_letter(columnasResp[sujeto]) + str(sesion + 3)] = mediaResPalForzDiscRef
                    esccolumnas(hojaind, value[3], value[5], respuestas, value[6])
                    # ¿Usar un diccionario en lugar de una tupla en analysisLis? Accesar a cada valor por su etiqueta y
                    # no por su posición.

            sujetoWb.save(dirConv + subList[subject] + suffix + str(session) + '.xlsx')
    workbook.save(dirConv + fileName)

# # Resumen
# # Revisar si el archivo de resumen ya existe. De lo contrario, crearlo.
# if archivo in listdir(directorioConvertidos):
#     print('Archivo encontrado. Abriendo...')
#     wb = load_workbook(directorioConvertidos + archivo)
# else:
#     print('Archivo no encontrado. Creando...')
#     wb = Workbook()
#
# # Crear todas las hojas.
# proporciones = hoja('Proporciones')
# respuestas = hoja('Respuestas')
# latencias = hoja('Latencias')
# comedero = hoja('Comedero')
# escapes = hoja('Escapes')
# latNosepoke = hoja('LatNosepoke')
# escapeForz = hoja('EscapesForzados')
# latEscForz = hoja('LatEscapeForz')
#
# # Loop principal.
# for sujeto in range(len(sujetos)):
#     print('\nIntentando sujeto ' + sujetos[sujeto] + '...')
#     for sesion in sesionesPresentes[sujeto]:
#         print('Intentando sesión ' + str(sesion) + '...')
#         sujetoWb = load_workbook(directorioConvertidos + sujetos[sujeto] + '_ESCAPE_' + str(sesion) + '.xlsx')
#         sujetoWs = sujetoWb.worksheets[0]
#         tiempo = sujetoWs['O']
#         marcadores = sujetoWs['P']
#
#         # Abrir o crear la hoja para pegar respuestas individuales.
#         if 'Respuestas por ensayo' not in sujetoWb.sheetnames:
#             hojaind = sujetoWb.create_sheet('Respuestas por ensayo')
#         else:
#             hojaind = sujetoWb['Respuestas por ensayo']
#
#         # Proporciones
#         proporciones[get_column_letter(columnasProp[sujeto]) + str(sesion + 3)] = sujetoWs.cell(14, 6).value
#         # Se toma el valor de la celda F14, que corresponde a los valores de fila 14 y columna 6.
#
#         # Conteo de respuestas en palancas en ensayos forzados.
#         # Se escriben las respuestas por ensayo en el archivo individual.
#         # Se resta una unidad a las medias debido a que conteoresp() cuenta también la respuesta que inicia el ensayo.
#         resPalForzDiscRef = conteoresp(114, 180, 202)
#         mediaResPalForzDiscRef = mean(resPalForzDiscRef) - 1
#         esccolumnas('PalForzDiscRef', 1, resPalForzDiscRef, True)
#         # 114: Inicio TL Forz Disc Ref  //  180: Fin ensayo  //  202: Resp Pal Disc
#         resPalForzDiscNoRef = conteoresp(115, 180, 202)
#         mediaResPalForzDiscNoRef = mean(resPalForzDiscNoRef) - 1
#         esccolumnas('PalForzDiscNoRef', 3, resPalForzDiscNoRef, True)
#         # 115: Inicio TL Forz Disc NoRef  //  180: Fin ensayo (por TF)  //  202: Res Pal Disc
#         resPalForzNoDisc1 = conteoresp(134, 180, 201)
#         mediaResPalForzNoDisc1 = mean(resPalForzNoDisc1) - 1
#         esccolumnas('PalForzNoDisc1', 5, resPalForzNoDisc1, True)
#         # 134: Inicio TL Forz NoDisc 1  //  180: Fin ensayo  //  201: Resp Pal NoDisc
#         resPalForzNoDisc2 = conteoresp(137, 180, 201)
#         mediaResPalForzNoDisc2 = mean(resPalForzNoDisc2) - 1
#         esccolumnas('PalForzNoDisc2', 7, resPalForzNoDisc2, True)
#         # 137: Inicio TL Forz NoDisc 2  //  180: Fin ensayo  //  201: Resp Pal NoDisc
#         respuestas[get_column_letter(columnasResp[sujeto]) + str(sesion + 3)] = mediaResPalForzDiscRef
#         respuestas[get_column_letter(columnasResp[sujeto] + 1) + str(sesion + 3)] = mediaResPalForzDiscNoRef
#         respuestas[get_column_letter(columnasResp[sujeto] + 2) + str(sesion + 3)] = mediaResPalForzNoDisc1
#         respuestas[get_column_letter(columnasResp[sujeto] + 3) + str(sesion + 3)] = mediaResPalForzNoDisc2
#
#         # Conteo de latencias a palancas en eslabones iniciales.
#         # Las latencias por ensayo se pegan en el archivo individual.
#         # La mediana no se escribe de inmediato. La lista de latencias por ensayo se almacena para sumarse a la lista
#         # de latencias a palanca en ensayos de escape forzado. La mediana se obtiene de la lista total.
#         latPalDisc = conteolat(112, 113)
#         # medianaLatPalDisc = median(latPalDisc)
#         esccolumnas('LatPalDisc', 9, latPalDisc, False)
#         # 112: IL Forz Disc  //  113: Res Pal
#         latPalNoDisc = conteolat(132, 133)
#         # medianaLatPalNoDisc = median(latPalNoDisc)
#         esccolumnas('LatPalNoDisc', 11, latPalNoDisc, False)
#         # 132: IL Forzado NoDisc  //  133: Res Pal
#         # latencias[get_column_letter(columnasLatPal[sujeto]) + str(sesion + 3)] = medianaLatPalDisc
#         # latencias[get_column_letter(columnasLatPal[sujeto] + 1) + str(sesion + 3)] = medianaLatPalNoDisc
#
#         # Conteo de respuestas en comederos en ensayos forzados
#         # Las respuestas por ensayo se pegan en el archivo individual.
#         resComForzDiscRef = conteoresp(114, 16, 203)
#         mediaResComForzDiscRef = mean(resComForzDiscRef)
#         esccolumnas('ComForzDiscRef', 13, resComForzDiscRef, False)
#         # 114: Inicio TL Forz Disc Ref  //  16: Fin ensayo  //  203: Res Com
#         resComForzDiscNoRef = conteoresp(115, 117, 203)
#         mediaResComForzDiscNoRef = mean(resComForzDiscNoRef)
#         esccolumnas('ComForzDiscNoRef', 15, resComForzDiscNoRef, False)
#         # 115: Inicio TL Forz Disc NoRef  //  117: Fin ensayo (por TF)  //  203: Res Com
#         resComForzNoDisc1 = conteoresp(134, 40, 203)
#         mediaResComForzNoDisc1 = mean(resComForzNoDisc1)
#         esccolumnas('ComForzNoDisc1', 17, resComForzNoDisc1, False)
#         # 134: Inicio TL Forz NoDisc 1  //  40: Fin ensayo  //  203: Res Com
#         resComForzNoDisc2 = conteoresp(137, 43, 203)
#         mediaResComForzNoDisc2 = mean(resComForzNoDisc2)
#         esccolumnas('ComForzNoDisc2', 19, resComForzNoDisc2, False)
#         # 137: Inicio TL Forz NoDisc 2  //  43: Fin ensayo  //  203: Res Com
#         comedero[get_column_letter(columnasResp[sujeto]) + str(sesion + 3)] = mediaResComForzDiscRef
#         comedero[get_column_letter(columnasResp[sujeto] + 1) + str(sesion + 3)] = mediaResComForzDiscNoRef
#         comedero[get_column_letter(columnasResp[sujeto] + 2) + str(sesion + 3)] = mediaResComForzNoDisc1
#         comedero[get_column_letter(columnasResp[sujeto] + 3) + str(sesion + 3)] = mediaResComForzNoDisc2
#
#         # Conteo de respuestas en nosepoke.
#         escapeForzDiscRef = conteototal(301)
#         escapeForzDiscNoRef = conteototal(302)
#         escapeForzNoDisc1 = conteototal(303)
#         escapeForzNoDisc2 = conteototal(304)
#         escapeLibDiscRef = conteototal(305)
#         escapeLibDiscNoRef = conteototal(306)
#         escapeLibNoDisc1 = conteototal(307)
#         escapeLibNoDisc2 = conteototal(308)
#         # 301 - 308: Respuestas nosepoke por tipo de ensayo.
#         escapes[get_column_letter(columnasEscapes[sujeto]) + str(sesion + 3)] = escapeForzDiscRef
#         escapes[get_column_letter(columnasEscapes[sujeto] + 1) + str(sesion + 3)] = escapeForzDiscNoRef
#         escapes[get_column_letter(columnasEscapes[sujeto] + 2) + str(sesion + 3)] = escapeForzNoDisc1
#         escapes[get_column_letter(columnasEscapes[sujeto] + 3) + str(sesion + 3)] = escapeForzNoDisc2
#         escapes[get_column_letter(columnasEscapes[sujeto] + 4) + str(sesion + 3)] = escapeLibDiscRef
#         escapes[get_column_letter(columnasEscapes[sujeto] + 5) + str(sesion + 3)] = escapeLibDiscNoRef
#         escapes[get_column_letter(columnasEscapes[sujeto] + 6) + str(sesion + 3)] = escapeLibNoDisc1
#         escapes[get_column_letter(columnasEscapes[sujeto] + 7) + str(sesion + 3)] = escapeLibNoDisc2
#
#         # Conteo de latencias a nosepoke.
#         # Las latencias por ensayo se pegan en el archivo individual.
#         latEscForzDiscRef = conteolat(114, 301)
#         medianaLatEscForzDiscRef = median(latEscForzDiscRef)
#         esccolumnas('LatEscForzDiscRef', 21, latEscForzDiscRef, False)
#
#         latEscForzDiscNoRef = conteolat(115, 302)
#         medianaLatEscForzDiscNoRef = median(latEscForzDiscNoRef)
#         esccolumnas('LatEscForzDiscNoRef', 23, latEscForzDiscNoRef, False)
#
#         latEscForzNoDisc1 = conteolat(134, 303)
#         medianaLatEscForzNoDisc1 = median(latEscForzNoDisc1)
#         esccolumnas('LatEscForzNoDisc1', 25, latEscForzNoDisc1, False)
#
#         latEscForzNoDisc2 = conteolat(137, 304)
#         medianaLatEscForzNoDisc2 = median(latEscForzNoDisc2)
#         esccolumnas('LatEscForzNoDisc2', 27, latEscForzNoDisc2, False)
#
#         latEscLibDiscRef = conteolat(154, 305)
#         medianaLatEscLibDiscRef = median(latEscLibDiscRef)
#         esccolumnas('LatEscLibDiscRef', 29, latEscLibDiscRef, False)
#
#         latEscLibDiscNoRef = conteolat(155, 306)
#         medianaLatEscLibDiscNoRef = median(latEscLibDiscNoRef)
#         esccolumnas('LatEscLibDiscNoRef', 31, latEscLibDiscNoRef, False)
#
#         latEscLibNoDisc1 = conteolat(157, 307)
#         medianaLatEscLibNoDisc1 = median(latEscLibNoDisc1)
#         esccolumnas('LatEscLibNoDisc1', 33, latEscLibNoDisc1, False)
#
#         latEscLibNoDisc2 = conteolat(160, 308)
#         medianaLatEscLibNoDisc2 = median(latEscLibNoDisc2)
#         esccolumnas('LatEscLibNoDisc2', 35, latEscLibNoDisc2, False)
#         # El primer marcador es el inicio de su tipo de ensayo; el segundo, la respuesta de escape correspondiente.
#
#         latNosepoke[get_column_letter(columnasEscapes[sujeto]) + str(sesion + 3)] = medianaLatEscForzDiscRef
#         latNosepoke[get_column_letter(columnasEscapes[sujeto] + 1) + str(sesion + 3)] = medianaLatEscForzDiscNoRef
#         latNosepoke[get_column_letter(columnasEscapes[sujeto] + 2) + str(sesion + 3)] = medianaLatEscForzNoDisc1
#         latNosepoke[get_column_letter(columnasEscapes[sujeto] + 3) + str(sesion + 3)] = medianaLatEscForzNoDisc2
#         latNosepoke[get_column_letter(columnasEscapes[sujeto] + 4) + str(sesion + 3)] = medianaLatEscLibDiscRef
#         latNosepoke[get_column_letter(columnasEscapes[sujeto] + 5) + str(sesion + 3)] = medianaLatEscLibDiscNoRef
#         latNosepoke[get_column_letter(columnasEscapes[sujeto] + 6) + str(sesion + 3)] = medianaLatEscLibNoDisc1
#         latNosepoke[get_column_letter(columnasEscapes[sujeto] + 7) + str(sesion + 3)] = medianaLatEscLibNoDisc2
#
#         # Conteo de escapes forzados
#         escForzDisc = conteototal(403)
#         escForzNoDisc = conteototal(406)
#         # 403 y 406: Respuestas en nosepoke en ensayos forzados de escape Disc y NoDisc
#         escapeForz[get_column_letter(columnasEscForz[sujeto]) + str(sesion + 3)] = escForzDisc
#         escapeForz[get_column_letter(columnasEscForz[sujeto] + 1) + str(sesion + 3)] = escForzNoDisc
#
#         # Latencias nosepoke escape forzado
#         latEscForzDisc = conteolat(401, 403)
#         medianaLatEscForzDisc = median(latEscForzDisc)
#         esccolumnas('LatEscForzDisc', 37, latEscForzDisc, False)
#
#         latEscForzNoDisc = conteolat(404, 406)
#         medianaLatEscForzNoDisc = median(latEscForzNoDisc)
#         esccolumnas('LatEscForzNoDisc', 39, latEscForzNoDisc, False)
#
#         latEscForz[get_column_letter(columnasLatPal[sujeto]) + str(sesion + 3)] = medianaLatEscForzDisc
#         latEscForz[get_column_letter(columnasLatPal[sujeto] + 1) + str(sesion + 3)] = medianaLatEscForzNoDisc
#
#         # Latencias palanca escape forzado
#         latPalEscForzDisc = conteolat(401, 402)
#         esccolumnas('LatPalEscForzDisc', 41, latPalEscForzDisc, False)
#
#         latPalEscForzNoDisc = conteolat(404, 405)
#         esccolumnas('LatPalEscForzNoDisc', 43, latPalEscForzNoDisc, False)
#
#         # Escribir latencia a palancas. Se incluyen latencias en ensayos forzados normales y forzados de escape.
#         latencias[get_column_letter(columnasLatPal[sujeto]) + str(sesion + 3)] = median(latPalDisc + latPalEscForzDisc)
#         latencias[get_column_letter(columnasLatPal[sujeto] + 1) + str(sesion + 3)] = median(latPalNoDisc + latPalEscForzNoDisc)
#
#         sujetoWb.save(directorioConvertidos + sujetos[sujeto] + '_ESCAPE_' + str(sesion) + '.xlsx')
#
# wb.save(directorioConvertidos + archivo)
