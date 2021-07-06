import pandas
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re

directorio = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Pruebas/'
sujetos = ['E3', 'E4', 'E5', 'E7', 'E8', 'E9']
sesionInicial = 1
sesionFinal = 3

for sujeto in range(len(sujetos)):
    for sesion in range(sesionInicial, sesionFinal + 1):
        print('Convirtiendo sesión ' + str(sesion) + ' de sujeto ' + sujetos[sujeto] + '.')
        # Primero se leen los metadatos y se escriben en el archivo convertido.
        encabezado = pandas.read_csv(directorio + sujetos[sujeto] + '_LIBRES_' + str(sesion), header=None, nrows=12)
        encabezado.to_excel(directorio + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx', startcol=0, index=False,
                            header=None)

        # Openpyxl lee ese archivo y almacena la columna que contiene los metadatos.
        encabezado = load_workbook(directorio + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx')
        hojaencabezado = encabezado.active
        columnaencabezado = []
        for col in hojaencabezado['A']:
            columnaencabezado.append(col.value)

        # Pandas escribe ahora el resto de los datos, pero al guardar sobreescribe los metadatos.
        datos = pandas.read_csv(directorio + sujetos[sujeto] + '_LIBRES_' + str(sesion), sep=r"\s+", skiprows=13,
                                header=None)
        datos.to_excel(directorio + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx', startcol=0, startrow=11,
                       index=False, header=None)

        # Para recuperarlos, openpyxl abre este nuevo archivo creado por pandas y escribe en él la columna que tenía
        # almacenada. Openpyxl no sobreescribe.
        archivoCompleto = load_workbook(directorio + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx')
        hojaCompleta = archivoCompleto.active
        for fila in range(1, len(columnaencabezado)):
            hojaCompleta['A' + str(fila)] = columnaencabezado[fila - 1]
        archivoCompleto.save(directorio + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx')

        # Generar una lista que contenga sub-listas con todos los valores de las listas dadas por Med.
        # Funciona para cualquier cantidad de listas.
        metalista = [[]]
        contadormetalista = 0
        columna1 = hojaCompleta['B']
        columna2 = hojaCompleta['C']
        columna3 = hojaCompleta['D']
        columna4 = hojaCompleta['E']
        columna5 = hojaCompleta['F']
        for i in range(len(columna1)):
            if columna1[i].value is not None:
                metalista[contadormetalista].append(str(columna1[i].value))
                if columna2[i].value is not None:
                    metalista[contadormetalista].append(str(columna2[i].value))
                    if columna3[i].value is not None:
                        metalista[contadormetalista].append(str(columna3[i].value))
                        if columna4[i].value is not None:
                            metalista[contadormetalista].append(str(columna4[i].value))
                            if columna5[i].value is not None:
                                metalista[contadormetalista].append(str(columna5[i].value))

            elif columna1[i].value is None and i > 11:
                metalista.append([])
                contadormetalista += 1

        # Escribir cada sub-lista en una columna de excel (aun no se separa por punto decimal).
        # Se utilizan expresiones regulares (regex) para indicar al programa que debe añadir ceros cuando pandas
        # los ha eliminado (cuando están al final de una cifra después de un punto decimal).
        regex1 = re.compile(r'^\d{1,5}\.\d{2}$')
        for i in range(len(metalista)):
            for j in range(len(metalista[i])):
                if regex1.search(metalista[i][j]):
                    metalista[i][j] += '0'
                hojaCompleta[get_column_letter((i*2)+9) + str(j + 1)] = metalista[i][j]
        archivoCompleto.save(directorio + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx')
        # Separar columnas por punto decimal.
        separable1 = []
        separable2 = []
        for col in hojaCompleta['M']:
            if col.value is not None:
                separable1.append(str(col.value).split('.')[0])
                separable2.append(str(col.value).split('.')[1])
        separable3 = []
        separable4 = []
        for col in hojaCompleta['O']:
            if col.value is not None:
                separable3.append(str(col.value).split('.')[0])
                separable4.append(str(col.value).split('.')[1])

        # Pegar de nuevo las columnas
        for i in range(len(separable1)):
            hojaCompleta['M' + str(i + 1)] = int(separable1[i])
            hojaCompleta['N' + str(i + 1)] = int(separable2[i])
        for i in range(len(separable3)):
            hojaCompleta['O' + str(i + 1)] = int(separable3[i])
            hojaCompleta['P' + str(i + 1)] = int(separable4[i])

        archivoCompleto.save(directorio + sujetos[sujeto] + '_LIBRES_' + str(sesion) + '.xlsx')
print('Hecho.')
