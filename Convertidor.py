import pandas
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re

directorioBrutos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/PruebasBrutos/'
directorioConvertidos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Flex/'
sujetos = ['E3', 'E4', 'E5', 'E7', 'E8', 'E9']
sesionInicial = 1
sesionFinal = 3

for sesion in range(sesionInicial, sesionFinal + 1):
    for sujeto in range(len(sujetos)):
        print('Convirtiendo sesión ' + str(sesion) + ' de sujeto ' + sujetos[sujeto] + '.')
        # Pandas lee los datos y los escribe en el archivo convertido en 6 columnas separando por los espacios.
        # El argumento names indica cuántas columnas se crearán. Evita errores cuando se edita el archivo de Med.
        datos = pandas.read_csv(directorioBrutos + sujetos[sujeto] + '_LIBRES_' + str(sesion), header=None,
                                names=range(6), sep=r'\s+')
        datos.to_excel(directorioConvertidos + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx',
                       index=False, header=None)

        # Openpyxl abre el archivo creado por pandas, lee la hoja y la almacena en la variable hojaCompleta.
        archivoCompleto = load_workbook(directorioConvertidos + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx')
        hojaCompleta = archivoCompleto.active

        # Se genera una lista que contenga sub-listas con todos los valores de las listas dadas por Med.
        # Funciona para cualquier cantidad de listas.
        # Los datos se convierten en flotantes para que todos tengan punto decimal, y luego en string para que mas
        # adelante el método split los pueda separar por el punto.
        metalista = [[]]
        contadormetalista = 0
        columna1 = hojaCompleta['B']
        columna2 = hojaCompleta['C']
        columna3 = hojaCompleta['D']
        columna4 = hojaCompleta['E']
        columna5 = hojaCompleta['F']
        for i in range(11, len(columna1)):
            if columna1[i].value is not None:
                metalista[contadormetalista].append(str(float(columna1[i].value)))
                if columna2[i].value is not None:
                    metalista[contadormetalista].append(str(float(columna2[i].value)))
                    if columna3[i].value is not None:
                        metalista[contadormetalista].append(str(float(columna3[i].value)))
                        if columna4[i].value is not None:
                            metalista[contadormetalista].append(str(float(columna4[i].value)))
                            if columna5[i].value is not None:
                                metalista[contadormetalista].append(str(float(columna5[i].value)))
            else:
                metalista.append([])
                contadormetalista += 1

        # Escribir cada sub-lista en una columna de excel separando por punto decimal.
        # Se utilizan expresiones regulares (regex) para indicar al programa que debe añadir ceros cuando pandas
        # los ha eliminado (cuando están al final de una cifra después de un punto decimal).
        regex1 = re.compile(r'^\d+\.\d{2}$')
        for i in range(len(metalista)):
            for j in range(len(metalista[i])):
                if regex1.search(metalista[i][j]):
                    metalista[i][j] += '0'
                hojaCompleta[get_column_letter((i * 2) + 9) + str(j + 1)] = int(metalista[i][j].split('.')[0])
                hojaCompleta[get_column_letter((i * 2) + 10) + str(j + 1)] = int(metalista[i][j].split('.')[1])
        archivoCompleto.save(directorioConvertidos + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx')
    print('\n')
