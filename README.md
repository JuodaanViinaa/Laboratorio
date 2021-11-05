# Laboratorio
Código utilizado en el laboratorio 101.

Estos _scripts_ se sirven de las librerías [Openpyxl](https://openpyxl.readthedocs.io/en/stable/index.html) y [Pandas](https://pandas.pydata.org/pandas-docs/stable/), por lo que será útil leer su documentación para entender algunas de las funciones utilizadas.

## Librería _Funciones.py_

Esta librería provee una manera simple de analizar los datos entregados por el programa de MedPC. Contiene funciones útiles para tareas y análisis básicos. 

De manera general la librería permite escanear una carpeta en la que se encuentran los archivos de texto sin formato entregados por MedPC que se desean analizar. Con base en ellos determina los sujetos y sesiones por analizar, convierte los archivos a formato ".xlsx" y separa las listas crudas en columnas más legibles, realiza los conteos de respuestas, latencias, o distribuciones de respuesta que el usuario declare, y finalmente escribe los resultados en archivos individuales para cada sujeto y en un archivo de resumen. Tras declarar todas las variables pertinentes el análisis completo de uno o más días de experimentos puede realizarse con un clic.

Sin embargo, la librería requiere de la declaración de variables específicas y su llamada en forma de argumentos en las funciones pertinentes.

Las variables necesarias para utilizar la librería son:

* Una variable que contenga el nombre del archivo de resumen en forma de _string_ en que se guardarán los datos. El _string_ debe contener la extensión ".xlsx". Ejemplo:
```python
archivo_de_resumen = "Igualacion.xlsx"
```

* Una variable que contenga la [dirección absoulta](https://www.geeksforgeeks.org/absolute-relative-pathnames-unix/) del directorio temporal en que se almacenarán los datos brutos antes de su análisis. Es importante que el último caracter del _string_ sea una diagonal `/`. Ejemplo:
```python
directorioTemporal = "C:/Users/Admin/Desktop/Direccion/De/Tu/Carpeta/DirectorioBrutos/"  # En el caso de Windows
directorioTemporal = "/home/usuario/Documents/Direccion/De/Tu/Carpeta/DirectorioBrutos/"  # En el caso de Unix
```

* Una variable que contenga la dirección absoluta del directorio permanente en que se guardarán los datos brutos después de haber sido utilizados (no es necesario mover los archivos manualmente después de su utilización; el programa se encarga de eso automáticamente). Ejemplo:
```python
directorioBrutos = "C:/Users/Admin/Desktop/Direccion/De/Tu/Carpeta/DirectorioTemporal/"  # En el caso de Windows
directorioBrutos = "/home/usuario/Documents/Direccion/De/Tu/Carpeta/DirectorioTemporal/"  # En el caso de Unix
```

* Una variable que contenga la dirección absoluta del directorio en que se guardarán los datos convertidos después del análisis. En este directorio se almacenarán tanto los archivos individuales con extensión ".xlsx" como el archivo de resumen. Ejemplo:
```python
directorioConvertidos = "C:/Users/Admin/Desktop/Direccion/De/Tu/Carpeta/DirectorioConvertidos/"  # En el caso de Windows
directorioConvertidos = "/home/usuario/Documents/Direccion/De/Tu/Carpeta/DirectorioConvertidos/"  # En el caso de Unix
```

* Una lista con los nombres de los sujetos en forma de _string_. Ejemplo:
```python
sujetos = ["Rata1", "Rata2", "Rata3", "Rata4"]
```

* Uno o más diccionarios que relacionen a cada sujeto con la columna en que sus datos se escribirán en cada hoja del archivo de resumen. Es decir: un mismo sujeto puede tener asociadas múltiples medidas (e.g., respuestas en palancas, respuestas en nosepoke, latencias, etc). Además, medidas distintas pueden tener subdivisiones diferentes (e.g., puede haber dos medidas de respuestas a palancas (izquierda y derecha), y una sola medida para respuestas a nosepoke). Así, es posible que en una hoja se encuentre un formato similar a este:

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;![image](https://user-images.githubusercontent.com/87039101/140447057-87d56167-8fe5-4322-97eb-50f3902f9b95.png)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Mientras que en otra se puede encontrar un formato similar a este: 

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;![image](https://user-images.githubusercontent.com/87039101/140447104-a501e12a-658e-4b7a-929e-46dc8c886edf.png)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Como se puede ver, medidas distintas para un mismo sujeto requerirían de cantidades distintas de columnas. Por ello será necesario declarar al menos dos diccionarios: uno que relacione a los sujetos con el espacio que ocupan en la primera hoja, y otro que los relacione con el espacio que ocupan en la segunda hoja. Estos diccionarios solamente necesitan declarar la primera columna ocupada. El resto es manejado más adelante. Así, un ejemplo de diccionarios sería:
```python
columnas_palancas = {"Rata1": 1, "Rata2": 4, "Rata3": 7,}
columnas_nosepoke = {"Rata1": 1, "Rata2": 3, "Rata3": 5,}
```

* Tres listas vacías que serán pobladas por las propias funciones de la librería y que contendrán las sesiones por analizar para cada sujeto, los marcadores, y el tiempo en vigésimas de segundo. La primera lista será poblada por la función `purgeSessions()`. La segunda y tercera serán pobladas durante el análisis principal de la función `analyze()`. Ejemplo:
```python
sesiones_presentes = []
marcadores = []
tiempo = []
```

* Finalmente, el corazón de la librería es una lista de diccionarios que dictará las medidas que serán extraídas de los marcadores y del tiempo real, al igual que la manera de escribirlas en los archivos individuales y de resumen. Esta lista tiene un formato específico que puede obtenerse ejecutando la función `template()` incluida con la librería. Cada diccionario de la lista declara la función a utilizar (contar respuestas por ensayo, contar respuestas totales, contar latencias, contar respuestas por _bin_ de tiempo), junto con sus parámetros pertinentes (marcadores, columna en que se escribirán los datos en los archivos individuales y de resumen, etc). Ejemplo:
```python
analysis_list = [
    {"conteoresp": {"mark1": 111, "mark2": 222, "mark3": 333,
                    "substract": True,
                    "column": 1,
                    "header": "Palanca Izq",
                    "sheet": "Palancas",
                    "summary_column_list": columnas_palancas,
                    "offset": 0,
                    }},
    {"conteoresp": {"mark1": 444, "mark2": 555, "mark3": 666,
                    "substract": True,
                    "column": 2,
                    "header": "Palanca Der",
                    "sheet": "Palancas",
                    "summary_column_list": columnas_palancas,
                    "offset": 1,
                    }},
    {"conteototal": {"mark1": 777,
                     "column": 3,
                     "header": "Nosepoke",
                     "sheet": "Nosepoke",
                     "summary_column_list": columnas_nosepoke,
                     "offset": 0,
                     }},
    ]
```

****
****


De manera más específica, hay cinco funciones utilizables y, por lo tanto, cinco formatos de diccionario. Son los siguientes:

### Fetch
```python
analysis_list = [
{"fetch": {"cell_row": 10,
           "cell_column": 10,
           "sheet": "Sheet_1",
           "summary_column_list": column_dictionary,
           "offset": 0
           }}
]
```

Esta función permite extraer directamente un único dato de los archivos ".xlsx" individuales. Es funcional, por ejemplo, para extraer rápidamente el número de ensayos completados (si es que éste se encuentra dentro de alguna de las listas otorgadas por MedPC). Los argumentos `"cell_row"` y `"cell_column"` dictan la posición de la celda que se quiere extraer: un dato que se encuentra, por ejemplo, en la celda "B7" requerirá de los argumentos `"cell_row": 7` y `"cell_column": 2`. El argumento `"sheet"` indica el nombre de la hoja del archivo de resumen en que se escribirá el dato extraído. 

El argumento `"summary_column_list"` indica el diccionario que asocia a cada sujeto con la columna particular en que se escribirá su dato. 

El argumento `"offset"` es un caso especial: en ocasiones se requerirá que medidas similares de un mismo sujeto sean escritas en columnas adyacentes de una misma hoja. Por ejemplo:

![image](https://user-images.githubusercontent.com/87039101/140452655-3109ae8b-3256-4596-93fe-d0b78196fd59.png)

En casos así no es necesario declarar diccionarios adicionales que asocien cada medida con una columna específica. Una manera más económica es declarar en varias medidas un único diccionario que asocie al sujeto con una columna "base", y después agregar incrementalmente unidades al argumento `"offset"`. Cada unidad en `"offset"` desplazará la medida en cuestión una columna hacia la derecha. Por ejemplo:
```python
analysis_list = [
{"fetch": {"cell_row": 10,
           "cell_column": 10,
           "sheet": "Sheet_1",
           "summary_column_list": column_dictionary,
           "offset": 0
           }},
{"fetch": {"cell_row": 20,
           "cell_column": 20,
           "sheet": "Sheet_1",
           "summary_column_list": column_dictionary,
           "offset": 1
           }},
{"fetch": {"cell_row": 30,
           "cell_column": 30,
           "sheet": "Sheet_1",
           "summary_column_list": column_dictionary,
           "offset": 2
           }},
]
```

Esto resultaría en tres columnas: la primera se encontraría en la posición declarada por el diccionario `column_dictionary`, mientras que las otras dos se encontrarían una y dos posiciones a la derecha.

### Conteoresp

```python
analysis_list = [
    {"conteoresp": {"measures": 2, # Opcional
                    "mark1": 111, "mark2": 222, "mark3": 333,
                    "mark4": 444, "mark5": 555, "mark6": 666, # Opcional
                    "substract": True, # Opcional
                    "column": 1,
                    "header": "Generic_title",
                    "sheet": "Sheet_2",
                    "summary_column_list": column_dictionary2,
                    "offset": 0,
                    }},
]
```

Esta función permite contar la cantidad de respuestas que ocurren entre el inicio y el fin de un tipo de ensayo particular. Una lista con todas las respuestas por ensayo se escribe en el archivo individual ".xlsx", y la media de la sesión se escribe en el archivo de resumen.

Los argumentos `"mark1"`, `"mark2"`, y `"mark3"` son los marcadores de inicio de ensayo, fin de ensayo, y respuesta de interés, respectivamente. 

El argumento `"substract"` es un argumento opcional que contempla el caso en el cual la respuesta que interesa contar sea también la respuesta que le da inicio al ensayo. En tal situación contar esa respuesta adicional resultaría en una imprecisión consistente que sobrestimaría en una unidad la cantidad total de respuestas por ensayo. Para solucionar esa situación, añadir el argumento `"substract": True` resultará en la resta de una unidad a cada uno de los conteos de respuestas por ensayo, lo cual nos devolverá a un conteo exacto.

Los argumentos `"column"` y `"header"` determinan la manera en que la lista completa de respuestas por ensayo se escribirá en el archivo individual ".xlsx". `"column"` indica la columna en la cual se pegará la lista (siendo que 1 = A, 2 = B, 3 = C, etc). El argumento `"header"` indica el rótulo que tendrá esa columna en su primera celda. 

Los argumentos `"sheet"` y `"summary_column_list"` determinan la manera en que la media de respuestas por ensayo de la sesión se escribirá en el archivo de resumen. El argumento `"sheet"` señala el nombre de la hoja de cálculo en que se escribirá el dato. Mientras, `"summary_column_list"` será el diccionario que asocia a sujetos con columnas (explicado anteriormente).

Esta función, junto con las funciones `conteolat` y `conteototal`, incorpora la posibilidad de realizar medidas "agregadas" o múltiples mediante el argumento `"measures"`: en algunas ocasiones es ventajoso sumar en una sola medida las respuesas (o latencias) provenientes de dos fuentes distintas. Como ejemplo se puede pensar en un caso en el cual haya respuestas en una palanca que lleven probabilísticamente a dos consecuencias diferentes y que, por descuido o planeación, tengan marcadores distintos. Las respuestas en ese caso deberán sumarse y contribuir a la misma media en el archivo de resumen. Para casos como ese el argumento `"measures"` permite agregar dentro de una misma medida fuentes distintas de información. `"measures"` indicará la cantidad de fuentes que se deberán agregar en la misma medida. Para cada medida adicional se deberán declarar además los marcadores pertinentes siguiendo la numeración lógica. Por ejemplo, para tres fuentes agregadas en una misma medida los argumentos serían:
```python
analysis_list = [
    {"conteoresp": {"measures": 2, # Opcional
                    "mark1": 123, "mark2": 124, "mark3": 125,
                    "mark4": 223, "mark5": 224, "mark6": 225, 
                    "mark7": 323, "mark8": 324, "mark9": 325, 
```

La atención debe centrarse en los dígitos que siguen a la palabra `mark`, notando que la numeración es consecutiva. Esta función puede manejar una cantidad indefinida de fuentes aglomeradas en una misma medida.

### Conteototal

```python
 analysis_list = [
    {"conteototal": {"measures": 2, # Opcional
                     "mark1": 111,
                     "mark2": 222, # Opcional
                     "column": 3,
                     "header": "Generic_title",
                     "sheet": "Sheet_4",
                     "summary_column_list": column_dictionary4,
                     "offset": 0,
                     }},
]
```

Esta función permite contar la cantidad total de respuestas en toda una sesión sin diferenciar entre ensayos. Sus argumentos son idénticos a los de `conteoresp()` salvo por dos excepciones: tiene un único argumento para marcador (el marcador de la respuesta de interés), por lo que al agregar más de una fuente en la misma medida la numeración saltará de uno en uno; y no tiene el argumento de `"substract"` en tanto que no hay una respuesta extra por descartar.

### Conteolat

```python
analysis_list = [
    {"conteolat": {"measures": 2, # Opcional
                   "mark1": 111, "mark2": 222,
                   "mark3": 333, "mark4": 444, # Opcional
                   "column": 2,
                   "header": "Generic_title",
                   "sheet": "Sheet_3",
                   "summary_column_list": column_dictionary3,
                   "offset": 0,
                   }},
]
```

## Convertidor.py

El _script_ del convertidor toma los archivos de Med en texto sin formato y los convierte en archivos con extensión .xlsx separando los datos en columnas con base en los espacios en blanco. Cada lista dada por Med es separada en dos columnas con base en el punto decimal. El tiempo en vigésimas de segundo y los marcadores se escriben en las columnas O y P, respectivamente. El convertidor escanea una carpeta temporal y con base en ella determina los sujetos faltantes y las sesiones a convertir. Además, es flexible y puede lidiar con cualquier número de listas y distintas organizaciones por columnas en los archivos de Med. Después de ser convertidos, los archivos brutos son movidos a una carpeta permanente.

El convertidor es una función con dos argumentos opcionales: 
```python
convertir(columnas, subfijo)
```

El argumento de `columnas` indica en cuántas columnas está dividido el archivo original dado por Med (el valor por defecto es `6`). Por ejemplo:

![image](https://user-images.githubusercontent.com/87039101/125010780-d31ce080-e02c-11eb-90cd-669ea14f8ab6.png)

En este caso los datos se dividen en cuatro columnas (incluyendo a la primera, que avanza de cinco en cinco), de modo que el argumento `columnas` será `4`.

El argumento `subfijo` indica el formato particular que tiene el nombre de los archivos a convertir (el valor por defecto es un _string_ vacío `''`). Por ejemplo, si el nombre de un archivo es "SujetoX_ALTER_1", el argumento `subfijo` será `'_ALTER_'` (con comillas, dado que se trata de un dato tipo _string_).

Para convertir un conjunto de archivos llamados "Sujeto1_ALTER_1", "Sujeto2_ALTER_1",..., "SujetoN_ALTER_1", cuyos archivos de Med están organizados en seis columnas (la organización por defecto dada por Med si no se declara explícitamente algo distinto), la función del convertidor deberá ser llamada de esta forma:

```python
convertir(subfijo='_ALTER_')
```

Para convertir un conjunto de archivos llamados "Sujeto1", "Sujeto2",..., "SujetoN", cuyos archivos de Med están organizados en dos columnas, la función del convertidor deberá ser llamada de esta forma:

```python
convertir(2)
```
o, alternativamente,

```python
convertir(columnas=2)
```

Si la organización de los archivos es en seis columnas y no se requiere agregar un 'subfijo' al nombre de los archivos, bastará llamar la función sin ningún argumento.
____
El código del convertidor presupone la existencia de las siguientes variables:

* `directorioBrutos`, que indicará la dirección de la carpeta en la cual se encuentran los archivos a convertir. Se deben utilizar diagonales hacia adelante ('/') y no hacia atrás ('\\'), como suele hacer Windows, para separar las carpetas en la dirección, y el último caracter de la dirección debe ser una diagonal hacia adelante.
* `directorioConvertidos`, que indicará la dirección de la carpeta en la cual se guardarán los archivos ya convertidos (además del archivo de resumen, si es que se utiliza el _script_ de ResumenUltimate.py). También será necesario usar diagonales hacia adelante y terminar con una diagonal.
* `sesionesPresentes`, una lista vacía que se llenará con las sesiones a analizar de cada sujeto.
* `sujetos`, que contendrá una lista con los nombres de los sujetos que forman el grupo a convertir.

Estas variables son las mismas utilizadas por el _script_ de resumen, de modo que solo es necesario declararlas una vez.

## ResumenUltimate.py

El _script_ está adaptado para el experimento de _escape_. 

Inicialmente, el código revisa una carpeta que contendrá temporalmente los archivos brutos a convertir y analizar. Busca los nombres de los sujetos declarados en la lista `sujetos` dentro de esta carpeta, y si algún sujeto no tiene ningún dato asociado, lo agrega a la lista `sujetosFaltantes`. Después se genera una lista hecha de sub-listas con las sesiones presentes para cada sujeto. Tener sesiones salteadas no debe generar ningún problema.

Se compara la lista `sujetosFaltantes` con la lista `sujetos`. Todos aquellos sujetos que se encuentren en la lista `sujetosFaltantes` son eliminados junto con sus columnas asociadas en sus listas respectivas.

El siguiente paso es la conversión de los archivos a formato .xlsx. Después de convertir los archivos, el código los lee nuevamente y realiza un conteo de respuestas y latencias que escribe en dos lugares: en una hoja nueva del archivo individual generado por el convertidor (donde se incluyen respuestas y latencias por ensayo) y en un archivo de resumen con extensión .xlsx (donde solo se incluyen medias para las respuestas y medianas para las latencias).

El _script_ de resumen declara funciones para las tareas repetitivas:
* `hoja(nombre)`: crea hojas de cálculo en el archivo de resumen con el nombre dado como argumento. Si una hoja con ese nombre ya existe, no se crea una hoja nueva, sino que ésta simplemente se abre. Esta función debe ser asignada a una variable (e.g., `latencias = hoja('Latencias')`).
* `conteoresp(inicioEnsayo, finEnsayo, respuesta)`: cuenta respuestas por tipo de ensayo. Los argumentos son los marcadores de Med para inicio de ensayo, fin de ensayo, y respuesta de interés. Esta función resulta en una lista que contiene la cantidad de ocasiones que la respuesta ocurrió entre cada inicio y fin de ensayo. Si no ocurrieron ensayos del tipo de interés, la lista resultante tendrá un único cero para poder hacer cálculos de medias y medianas. Es importante tener en cuenta que generalmente los ensayos son iniciados por una respuesta. Esta primera respuesta ocurre después del marcador de inicio de ensayo, pero no forma parte del conteo de respuestas real. Dependiendo de que la respuesta que inicia el ensayo sea también la respuesta de interés a contar puede ser necesario restar una unidad al calcular la media de respuestas por ensayo. Esta función debe ser asignada a una variable.
* `conteototal(respuesta)`: cuenta respuestas totales de un tipo dado en la sesión completa. El argumento es el marcador de Med de la respuesta de interés. Esta función resulta en un número entero que representa la cantidad de ocasiones que la respuesta ocurrió en la sesión completa. Esta función debe ser asignada a una variable.
* `conteolat(inicioEnsayo, respuesta)`: cuenta la latencia entre el inicio de un ensayo de tipo determinado y la primera respuesta de interés. Los argumentos son los marcadores de Med de inicio de ensayo y respuesta. La función resulta en una lista con la latencia de respuesta en cada ensayo. Esta función debe ser asignada a una variable.
* `esccolumnas(titulo, columna, lista, restar)`: escribe las respuestas y latencias por ensayo en el archivo convertido individual. El argumento de `titulo` es un _string_ e indica el encabezado que tendrá la columna en donde se escribirá la lista; `columna` es un número entero e indica la columna en la cual se escribirá la lista (1 = A, 2 = B, etc.); `lista` indicará el nombre de la variable que contiene la lista que será escrita en la columna (las listas serán los resultados de las funciones descritas anteriormente); y `restar` indicará si se debe restar una unidad a cada elemento de la lista antes de pegarlo en su columna individual (debido a que la función `conteoresp()` cuenta una respuesta adicional por ensayo en situaciones ya descritas), y puede tomar los valores de `True` y `False`. Si el argumento `restar` es `True` y hay un ensayo con cero respuestas, no se ejecuta la resta para evitar resultar en números negativos.

La sección principal del _script_ se basa en dos ciclos `for`: uno que cicla a través de sujetos y otro anidado en el primero que lo hace a través de sesiones.

El _script_ genera listas con los conteos de respuestas y latencias de cada sujeto, y escribe medias y medianas en los sitios correspondietes del archivo de resumen. Los sitios correspondientes son determinados consultando las listas `columnasProp`, `columnasResp`, `columnasLatPal`, `columnasEscapes`, y `columnasLatEsc`. Existe una lista por cada una de las variables que se analizan debido a que cada variable se escribe en una hoja distinta y ocupa una cantidad de columnas distinta. Cada lista declara el número correspondiente a la primera columna utilizada por cada uno de los sujetos, por lo que habrá tantos elementos en cada lista como sujetos en la lista de `sujetos`. 

## Resumen(xls).py

Este resumen está desacoplado del convertidor, y está hecho para utilizarse con el convertidor anterior del laboratorio (que produce archivos convertidos con extensión .xls en lugar de .xlsx). Tiene las mismas funciones que _ResumenUltimate.py_ salvo porque no pega datos por ensayo en el archivo individual. Es preferible no usar este script.
