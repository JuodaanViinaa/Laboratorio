import pandas
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re

directorioBrutos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/PruebasBrutos/'
directorioConvertidos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Flex/'
sujetos = ['E3', 'E4', 'E5', 'E7', 'E8', 'E9']
sesionInicial = 1
sesionFinal = 1


def convertir(columnas, subfijo):
    for sesion in range(sesionInicial, sesionFinal + 1):
        for sujeto in range(len(sujetos)):
            print('Convirtiendo sesión ' + str(sesion) + ' de sujeto ' + sujetos[sujeto] + '.')
            # Pandas lee los datos y los escribe en el archivo convertido en 6 columnas separando por los espacios.
            # El argumento names indica cuántas columnas se crearán. Evita errores cuando se edita el archivo de Med.
            datos = pandas.read_csv(directorioBrutos + sujetos[sujeto] + subfijo + str(sesion), header=None,
                                    names=range(columnas), sep=r'\s+')
            datos.to_excel(directorioConvertidos + sujetos[sujeto] + subfijo + str(sesion) + '.xlsx', index=False,
                           header=None)

            # Openpyxl abre el archivo creado por pandas, lee la hoja y la almacena en la variable hojaCompleta.
            archivoCompleto = load_workbook(directorioConvertidos + sujetos[sujeto] + subfijo + str(sesion) + '.xlsx')
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
            for i in range(len(metalista)):
                for j in range(len(metalista[i])):
                    if regex1.search(metalista[i][j]):
                        metalista[i][j] += '0'
                    hojaCompleta[get_column_letter((i * 2) + 9) + str(j + 1)] = int(metalista[i][j].split('.')[0])
                    hojaCompleta[get_column_letter((i * 2) + 10) + str(j + 1)] = int(metalista[i][j].split('.')[1])
            archivoCompleto.save(directorioConvertidos + sujetos[sujeto] + subfijo + str(sesion) + '.xlsx')
        print('\n')


convertir(6, '_LIBRES_')
