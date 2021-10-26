from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from os import listdir
from statistics import mean, median
import pandas
import re
from shutil import move


def purgeSessions(temporaryDirectory, subjectList, sessionList):
    """
    Se genera una lista temporal que contiene aquellos sujetos cuyos datos sí están en la carpeta temporal; después, se
    eliminan las columnas pertinentes de cada una de las listas de columnas. Los argumentos son: el directorio temporal
    en que están los datos brutos, una lista con los nombres de los sujetos, y las listas de columnas.\n
    :param temporaryDirectory: Directorio donde se almacenan temporalmente los datos brutos por analizar.
    :param subjectList: Lista con los nombres de todos los sujetos a analizar.
    :param sessionList: Lista vacía que contendrá las sesiones presentes por analizar para cada sujeto.
    :param columnLists: Listas con los valores de las columnas en que se pegarán los datos para analizar.
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
        print("All subjects have at least one session yet to analyze.")
    else:
        print(f"Missing subjects: {str(sujetosFaltantes)}")

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
        print(f"Sessions to analyze for subject {listaTemp[i]}: {sessionList[i]}")

    # Se eliminan los elementos pertinentes de las listas de columnas si algún sujeto falta.
    for sujetoFaltante in sujetosFaltantes:
        #     for columnList in columnLists:
        #         if sujetoFaltante in subjectList:
        #             del columnList[subjectList.index(sujetoFaltante)]
        del subjectList[subjectList.index(sujetoFaltante)]
    print("\n")


def convertir(dirTemp, dirPerm, dirConv, subjectList, presentSessions, columnas=6, subfijo='_', mover=True):
    """
    Convierte archivos de texto plano de MedPC en hojas de cálculo en formato *.xlsx.
    Separa cada lista en dos columnas con base en el punto decimal.\n
    :param dirTemp: Directorio donde se almacenan temporalmente los datos brutos por analizar.
    :param dirPerm: Directorio donde se almacenarán finalmente los datos brutos después de la conversión.
    :param dirConv: Directorio donde se almacenarán los archivos ya convertidos.
    :param subjectList: Lista con los nombres de todos los sujetos.
    :param presentSessions: Lista vacía rellenada por el programa con las sesiones presentes de cada sujeto.
    :param columnas: Cantidad de columnas en que están divididos los archivos de texto de Med. Valor por defecto: 6.
    :param subfijo: Identificador del nombre de los archivos. Ejemplo: '_Alter_'. Valor por defecto: ''.
    :param mover: Booleano que indica si los archivos deben ser movidos a la carpeta permanente tras su conversión.
    """
    for sjt in range(len(subjectList)):
        for ssn in presentSessions[sjt]:
            print(f"Converting session {str(ssn)} of subject {subjectList[sjt]}.")
            # Pandas lee los datos y los escribe en el archivo convertido en 6 columnas separando por los espacios.
            # El argumento names indica cuántas columnas se crearán. Evita errores cuando se edita el archivo de Med.
            datos = pandas.read_csv(dirTemp + subjectList[sjt] + subfijo + str(ssn), header=None,
                                    names=range(columnas), sep=r'\s+')
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
            for fila in range(12, len(columna1) + 1):
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
                    elif regex2.search(metalista[ii][j]):
                        metalista[ii][j] += '00'
                    hojaCompleta[get_column_letter((ii * 2) + columnas + 3) + str(j + 1)] = int(
                        metalista[ii][j].split('.')[0])
                    hojaCompleta[get_column_letter((ii * 2) + columnas + 4) + str(j + 1)] = int(
                        metalista[ii][j].split('.')[1])
            archivoCompleto.save(dirConv + subjectList[sjt] + subfijo + str(ssn) + '.xlsx')
            if mover:
                move(dirTemp + subjectList[sjt] + subfijo + str(ssn), dirPerm + subjectList[sjt] + subfijo + str(ssn))
        print('\n')


def createDocument(fileName, targetDirectory):
    """
    Se inspecciona el directorio objetivo en busca de un archivo de resumen. Si existe, el archivo es abierto. Si no,
    es creado. Esta función debe asignarse a una variable.\n
    :param fileName: Nombre del archivo de resumen.
    :param targetDirectory: Ubicación del directorio objetivo.
    :return: Objeto de clase Workbook (Openpyxl).
    """
    if fileName in listdir(targetDirectory):
        print('Summary file found. Opening...\n')
        wb = load_workbook(targetDirectory + fileName)
    else:
        print('Summary file not found. Creating...\n')
        wb = Workbook()
    return wb


def create_sheets(workbook, *sheets):
    """
    Se crea una lista con hojas de trabajo pertenecientes al directorio creado/abierto. Se pueden crear tantas listas
    como sea necesario. La lista tendrá el mismo orden que los argumentos de la función.\n
    :param workbook: Archivo de tipo Workbook (creado mediante Openpyxl) en el que se generarán las hojas de trabajo.
    :param sheets: Strings con los nombres que tendrá cada hoja de cálculo. La función admite una cantidad indefinida
    de hojas.
    :return: Diccionario con listas asociadas al nombre dado como parámetro.
    """
    sheet_dict = {}
    for sheet in sheets:
        if sheet not in workbook.sheetnames:
            new_sheet = workbook.create_sheet(sheet)
        else:
            new_sheet = workbook[sheet]
        sheet_dict[sheet] = new_sheet
    return sheet_dict


def fetch(sheet, origin_cell_row, origin_cell_column):
    """
    La función regresa el valor contenido dentro de una celda en una hoja de cálculo especificada.\n
    :param sheet: La hoja de cálculo leída.
    :param origin_cell_column: Columna en que se encuentra la celda.
    :param origin_cell_row: Fila en que se encuentra la celda.
    :return: Valor de la celda referida.
    """
    return sheet.cell(origin_cell_row, origin_cell_column).value


def conteoresp(marcadores, inicioEnsayo, finEnsayo, respuesta):  # Count_per_trial
    """
    Cuenta respuestas entre el marcador de inicio de ensayo y el de fin de ensayo. En caso de que el marcador de la
    respuesta a contar sea el mismo que el de la respuesta que da inicio al ensayo se contará una respuesta adicional.
    Esto puede corregirse en el análisis principal con la opción "substract", y en la función 'esccolumnas' mediante el
    parámetro "restar".\n
    :param marcadores: Lista con los marcadores.
    :param inicioEnsayo: Marcador de inicio de ensayo.
    :param finEnsayo: Marcador de fin de ensayo.
    :param respuesta: Marcador de respuesta a contar.
    :return: Lista con la cantidad de respuestas ocurridas por ensayo.
    """
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


def conteototal(marcadores, respuesta):
    """
    Cuenta la cantidad total de veces que ocurrió un marcador particular en la sesión independientemente de los ensayos.
    \n
    :param marcadores: Lista con los marcadores.
    :param respuesta: Marcador de respuesta a contar.
    :return: Integer con la cantidad de ocasiones que ocurrió una respuesta.
    """
    contador = 0
    for n in range(len(marcadores)):
        if marcadores[n].value == respuesta:
            contador += 1
    return contador


def conteolat(marcadores, tiempo, inicioensayo, respuesta):
    """
    Cuenta la latencia entre el inicio de un ensayo y una respuesta.\n
    :param marcadores: Lista con los marcadores.
    :param tiempo: Lista con el tiempo de la sesión.
    :param inicioensayo: Marcador de inicio de ensayo.
    :param respuesta: Marcador de respuesta de interés.
    :return: Lista con las latencias por ensayo.
    """
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


def esccolumnas(hojaind, titulo, columna, lista, restar):
    """
    Escribe listas completas en columnas. Útil para escribir los datos completos en los archivos individuales.\n
    :param hojaind: Objeto de tipo Worksheet (Openpyxl) en el que se escribirá la lista.
    :param titulo: Rótulo que se escribirá en la primera celda de la columna.
    :param columna: Número de la columna en que se escribirá la lista (1-A, 2-B, etc.).
    :param lista: Lista a escribir en la columna.
    :param restar: Booleano (True/False) que indica si debe restarse una unidad a cada uno de los valores de la lista.
    Deberá ser True si la respuesta contada es la misma que da inicio al ensayo y, por lo tanto, se lleva contada una
    respuesta adicional a las reales.
    """
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


def analyze(dirConv, fileName, subList, sessionList, suffix, workbook, sheetDict, analysisList, markColumn, timeColumn):
    """
    Función principal de análisis. Toma los conteos realizados por las otras funciones y pega medidas de tendencia
    central en un archivo general de resumen. Además pega las listas completas en los archivos convertidos
    individuales.\n
    :param dirConv: Dirección de la carpeta donde se guardarán los archivos convertidos y el archivo de resumen. La dirección debe ser absoluta y estar separada por diagonales hacia adelante.
    :param fileName: Nombre del archivo de resumen.
    :param subList: Lista con los nombres de los suejetos a analizar.
    :param sessionList: Lista con las sesiones presentes para cada sujeto. Inicialmente está vacía y es poblada por la función de purgeSessions
    :param suffix: Caracter o conjunto de caracteres que separa el nombre del sujeto del número de sesión en cada uno de los archivos. Ejemplo: "_"
    :param workbook: Hoja de cálculo generada por la función createDocument.
    :param sheetDict: Diccionario creado por la función create_sheets.
    :param analysisList: Lista de diccionarios con los parámetros para analizar cada magnitud.
    :param markColumn: Columna ocupada por la lista de marcadores en los archivos individuales. Para obtenerla es necesario correr solamente la función de convertir y revisar manualmente la columna en que están escritos los marcadores.
    :param timeColumn: Columna ocupada por la lista de tiempos en vigésimas de segundo. Debe ser revisada manualmente, también.
    """
    for subject in range(len(subList)):
        for session in sessionList[subject]:
            print(f"Trying session {session} of subject {subList[subject]}.")
            sujetoWb = load_workbook(dirConv + subList[subject] + suffix + str(session) + '.xlsx')
            sujetoWs = sujetoWb.worksheets[0]
            hojaind = sujetoWb.create_sheet('FullLists')
            tiempo = sujetoWs[timeColumn]
            marcadores = sujetoWs[markColumn]

            for analysis in analysisList:
                key, value = list(analysis.items())[0]

                if key == "conteoresp":
                    respuestas_totales = []
                    for mark_index in range(1, value.get("measures", 1) + 1):
                        respuesta_parcial = conteoresp(marcadores, value[f"mark{(mark_index * 3) - 2}"],
                                                       value[f"mark{(mark_index * 3) - 1}"],
                                                       value[f"mark{(mark_index * 3)}"])
                        respuestas_totales.extend(respuesta_parcial)
                    if value.get("substract", False):
                        respuestas_restadas = [resp - 1 if resp > 0 else resp for resp in respuestas_totales]
                        sheetDict[value["sheet"]][
                            get_column_letter(value["summary_column_list"][subList[subject]] + value["offset"]) + str(
                                session + 3)] = mean(respuestas_restadas)
                    else:
                        sheetDict[value["sheet"]][
                            get_column_letter(value["summary_column_list"][subList[subject]] + value["offset"]) + str(
                                session + 3)] = mean(respuestas_totales)
                    esccolumnas(hojaind, value["header"], value["column"], respuestas_totales,
                                value.get("substract", False))

                    # respuestas = conteoresp(marcadores, value["mark1"], value["mark2"], value["mark3"])
                    # if value.get("substract", False) and mean(respuestas) - 1 >= 0:
                    #     sheetDict[value["sheet"]][
                    #         get_column_letter(value["summary_column_list"][subList[subject]] + value["offset"]) + str(
                    #             session + 3)] = mean(respuestas) - 1
                    # else:
                    #     sheetDict[value["sheet"]][
                    #         get_column_letter(value["summary_column_list"][subList[subject]] + value["offset"]) + str(
                    #             session + 3)] = mean(respuestas)
                    # esccolumnas(hojaind, value["header"], value["column"], respuestas, value.get("substract", False))

                elif key == "conteolat":
                    latencias_totales = []
                    for mark_index in range(1, value.get("measures", 1) + 1):
                        latencia_parcial = conteolat(marcadores, tiempo, value[f"mark{(mark_index * 2) - 1}"],
                                                     value[f"mark{(mark_index * 2)}"])
                        latencias_totales.extend(latencia_parcial)
                    sheetDict[value["sheet"]][
                        get_column_letter(value["summary_column_list"][subList[subject]] + value["offset"]) + str(
                            session + 3)] = median(latencias_totales)
                    esccolumnas(hojaind, value["header"], value["column"], latencias_totales,
                                value.get("substract", False))
                    # latencias = conteolat(marcadores, tiempo, value["mark1"], value["mark2"])
                    # sheetDict[value["sheet"]][
                    #     get_column_letter(value["summary_column_list"][subList[subject]] + value["offset"]) + str(
                    #         session + 3)] = median(latencias)
                    # esccolumnas(hojaind, value["header"], value["column"], latencias, value.get("substract", False))

                elif key == "conteototal":
                    respuestas_totales = 0
                    for mark_index in range(1, value.get("measures", 1) + 1):
                        respuesta_parcial = conteototal(marcadores, value[f"mark{mark_index}"])
                        respuestas_totales += respuesta_parcial
                    sheetDict[value["sheet"]][
                        get_column_letter(value["summary_column_list"][subList[subject]] + value["offset"]) + str(
                            session + 3)] = respuestas_totales
                    esccolumnas(hojaind, value["header"], value["column"], [respuestas_totales],
                                value.get("substract", False))
                    # respuestasTot = str(conteototal(marcadores, value["mark1"]))
                    # sheetDict[value["sheet"]][
                    #     get_column_letter(value["summary_column_list"][subList[subject]] + value["offset"]) + str(
                    #         session + 3)] = respuestasTot
                    # esccolumnas(hojaind, value["header"], value["column"], [respuestasTot], value.get("substract", False))

                elif key == "fetch":
                    cell_value = fetch(sujetoWs, value["cell_row"], value["cell_column"])
                    sheetDict[value["sheet"]][
                        get_column_letter(value["summary_column_list"][subList[subject]] + value["offset"]) + str(
                            session + 3)] = cell_value

                # elif key == "agregate":
                #     key2, value2 = list(value.items())[0]
                #     if key2 == "conteolat":
                #         latencias_totales = []
                #         for mark_index in range(1, value2["measures"] + 1):
                #             latencia_parcial = conteolat(marcadores, tiempo, value2[f"mark{(mark_index * 2) - 1}"],
                #                                          value2[f"mark{(mark_index * 2)}"])
                #             latencias_totales.extend(latencia_parcial)
                #         sheetDict[value2["sheet"]][
                #             get_column_letter(value2["summary_column_list"][subList[subject]] + value2["offset"]) + str(
                #                 session + 3)] = median(latencias_totales)
                #         esccolumnas(hojaind, value2["header"], value2["column"], latencias_totales, value2.get("substract", False))
                #
                #     elif key2 == "conteoresp":
                #         respuestas_totales = []
                #         for mark_index in range(1, value2["measures"] + 1):
                #             respuesta_parcial = conteolat(marcadores, value2[f"mark{(mark_index * 3) - 2}"],
                #                                           value2[f"mark{(mark_index * 3) - 1}"],
                #                                           value2[f"mark{(mark_index * 3)}"])
                #             respuestas_totales.extend(respuesta_parcial)
                #         sheetDict[value2["sheet"]][
                #             get_column_letter(value2["summary_column_list"][subList[subject]] + value2["offset"]) + str(
                #                 session + 3)] = mean(respuestas_totales)
                #         esccolumnas(hojaind, value2["header"], value2["column"], respuestas_totales, value2.get("substract", False))
                #
                #     elif key2 == "conteototal":
                #         respuestas_totales = 0
                #         for mark_index in range(1, value2["measures"] + 1):
                #             respuesta_parcial = conteototal(marcadores, value2[f"mark{mark_index}"])
                #             respuestas_totales += respuesta_parcial
                #         sheetDict[value2["sheet"]][
                #             get_column_letter(value2["summary_column_list"][subList[subject]] + value2["offset"]) + str(
                #                 session + 3)] = respuestas_totales
                #         esccolumnas(hojaind, value2["header"], value2["column"], [respuestas_totales],
                #                     value2.get("substract", False))

            sujetoWb.save(dirConv + subList[subject] + suffix + str(session) + '.xlsx')
        print("\n")
    workbook.save(dirConv + fileName)


def template():
    print("""
    analysis_list = [
    {"fetch": {"sheet": "Sheet_1",
               "summary_column_list": columns,
               "cell_row": 10,
               "cell_column": 10,
               "offset": 0
               }},
    
    {"conteoresp": {"measures": 2, # Optional argument. Default value: 1
                    "mark1": 111, "mark2": 222, "mark3": 333,
                    "header": "Generic_title",
                    "sheet": "Sheet_2",
                    "column": 1,
                    "substract": True, # Optional argument. Default value: False
                    "summary_column_list": columnasResp,
                    "offset": 0}},
    
    {"conteolat": {"measures": 2, # Optional argument. Default value: 1
                   "mark1": 112, "mark2": 113,
                   "header": "LatPalDisc",
                   "sheet": "Latencias",
                   "column": 9,
                   "summary_column_list": columnasLatPal,
                   "offset": 0}},

    {"conteototal": {"measures": 2, # Optional argument. Default value: 1
                     "mark1": 301,
                     "header": "EscForzDiscRef",
                     "sheet": "Escapes",
                     "column": 21,
                     "summary_column_list": columnasEscapes,
                     "offset": 0}},
    ]
    """)
