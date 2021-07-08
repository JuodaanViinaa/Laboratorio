# Laboratorio
Código utilizado en el laboratorio 101


## Convertidor.py

El _script_ del convertidor toma los archivos de Med en texto sin formato y los convierte en archivos con extensión .xlsx separando los datos en columnas con base en los espacios en blanco. Cada lista dada por Med es separada en dos columnas con base en el punto decimal. El tiempo en vigésimas de segundo y los marcadores se pegan en las columnas O y P, respectivamente. El convertidor es flexible y funciona para cualquier cantidad de listas.

El convertidor es una función con dos argumentos: "columnas" y "subfijo".

El argumento de "columnas" debe indicar en cuántas columnas está dividido el archivo original dado por Med. Por ejemplo:

![image](https://user-images.githubusercontent.com/87039101/124951294-d38b8c00-dfd8-11eb-9d80-def3c77ae8c2.png)

En este caso los datos se dividen en seis columnas (contando a la primera, que avanza de cinco en cinco), de modo que el argumento "columnas" será "6" (sin comillas).

El argumento "subfijo" indica el formato particular que tiene el nombre de los archivos a convertir. Por ejemplo, si el nombre de los datos es "SujetoX_ALTER_1", el argumento "subfijo" será "'\_ALTER\_'" (con comillas, dado que se trata de un dato tipo _string_).

Para convertir un conjunto de archivos llamados "Sujeto1_ALTER_1", "Sujeto2_ALTER_1",..., "SujetoN_ALTER_1", cuyos archivos de Med están organizados en seis columnas (la organización por defecto dada por Med si no se declara explícitamente algo distinto), la función del convertidor deberá ser llamada de esta forma:

convertir(6, '\_ALTER\_')

El código del convertidor presupone la existencia de las siguientes variables:

* "directorioBrutos", que indicará la dirección de la carpeta en la cual se encuentran los archivos a convertir. Se deben utilizar diagonales hacia adelante ('/') y no hacia atrás ('\\'), como suele hacer Windows, para separar las carpetas de la dirección, y el último caracter de la dirección debe ser una diagonal hacia adelante.
* "directorioConvertidos", que indicará la dirección de la carpeta en la cual se guardarán los archivos ya convertidos (además del archivo de resumen, si es que se utiliza el _script_ de ResumenUltimate.py). También será necesario usar diagonales hacia adelante y terminar con una diagonal.
* "sesionInicial", que indica el número de la primera sesión a convertir.
* "sesionFinal", que indica el número de la última sesión a convertir.
* "sujetos", que contendrá una lista con los nombres de los sujetos que forman el grupo a convertir.

Estas variables son las mismas utilizadas por el _script_ de resumen, de modo que solo es necesario declararlas una vez.

## ResumenUltimate.py

El _script_ está adaptado para el experimento de _escape_ e integra al convertidor como su primer paso. Después de convertir los archivos, los lee nuevamente y realiza un conteo de respuestas y latencias que pega en dos lugares: en una hoja nueva del archivo individual generado por el convertidor (donde se incluyen respuestas y latencias por ensayo) y en un archivo de resumen con extensión .xlsx (donde solo se incluyen medias para las respuestas y medianas para las latencias).

## Resumen(xls).py

Este resumen está desacoplado del convertidor, y está hecho para utilizarse con el convertidor viejo del laboratorio (que produce archivos convertidos con extensión .xls en lugar de .xlsx). Tiene las mismas funciones que _ResumenUltimate.py_ salvo porque no pega datos por ensayo en el archivo individual. Es preferible no usar este script.
