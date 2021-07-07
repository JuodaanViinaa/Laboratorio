# Laboratorio
Código utilizado en el laboratorio 101


**Convertidor.py**

El _script_ del convertidor toma los archivos de Med en texto sin formato y los convierte en archivos con extensión .xlsx, separando los datos en columnas con base en los espacios.

**ResumenUltimate.py**

El _script_ está adaptado para el experimento de _escape_ e integra al convertidor como su primer paso. Después de convertir los archivos, los lee nuevamente y realiza un conteo de respuestas y latencias que pega en dos lugares: en una hoja nueva del archivo individual generado por el convertidor (donde se incluyen respuestas y latencias por ensayo) y en un archivo de resumen con extensión .xlsx (en donde solo se incluyen medias para las respuestas y medianas para las latencias).

**Resumen(xls).py**

Este resumen está desacoplado del convertidor, y está hecho para utilizarse con el convertidor viejo del laboratorio (que produce archivos convertidos con extensión .xls en lugar de .xlsx). Tiene las mismas funciones que _ResumenUltimate.py_ salvo porque no pega datos por ensayo en el archivo individual. Es preferible no usar este script.
