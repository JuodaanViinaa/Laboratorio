# MedPCPy

The purpose of this library is to provide an easy and accesible way to convert MedPC files to .xlsx (Excel, LibreOffice Calc) format; and then to extract and order the relevant data (response frecuencies, latencies, and distributions) without the need of much programming knowledge. After proper setup the entirety of the analysis of one or more days of experiments and one or more subjects can be done with a single click. The library delivers both individual files and a summary file: the individual files contain complete and properly labeled lists of all variables of interest; the summary file contains central tendency measures (either mean or median) for each variable.

This library uses functions from both [Openpyxl](https://openpyxl.readthedocs.io/en/stable/index.html) and [Pandas](https://pandas.pydata.org/pandas-docs/stable/). As such, it is advisable to be familiarized with them in order to understand the inner workings of some of its functions. It is, however, not necessary to know either of them to use this library.

## Quick start

Once the python script is properly set, an example workflow could be as follows:

1. Run experiments on a MedPC interface.
2. Transfer the raw MedPC files to a temporary directory which the script will read.
3. Run the python script. The script will automatically read the files, convert them, extract all declared measures, write them on individual files as well as on a summary file, and move the raw files to a permanent directory so that the temporary directory is empty once again.

No other interaction is needed as long as the files are properly located and named.
_____
The first step is to install the library with the command

```python
pip install medpcpy
```

and then import it to the working script with

```python
from medpcpy import *
```

to get access to all the necessary functions without the need to call `medpcpy.` on every use.

All of the work is performed by a single object of class `Analyzer` which contains methods to convert the MedPC files to .xlsx and then extract and summarize the relevant data. The `Analyzer` object requires several arguments to be initialized. These arguments are:

1. `fileName`, the name of the summary file.
2. `temporaryDirectory`, the directory in which raw MedPC files are stored before the analysis.
3. `permanentDirectory`, the directory to which raw MedPC files will be moved after analysis.
4. `convertedDirectory`, the directory in which individual .xlsx files and the summary file will be stored after the analysis.
5. `subjectList`, a list of _strings_ with the names of all subjects.
6. `suffix`, a _string_ which indicates the character or characters which separate the subject name from the session number in the raw MedPC filenames (ex.: if raw files are named "subject1_1", "subject2_1", etc., then the value for the `suffix` argument should be `"_"`).
	* The filenames must follow the format `"[subject name][spacing character][session number]"` so that the library can properly read them. Ex.: `"Rat1_pretraining_1"`, where `"_pretraining_"` is the spacing character and, thus, the value for the `suffix` argument.
7. `sheets`, a list of _strings_ which represent the names of each individual sheet which will be created in the summary file.
8. `analysisList`, a list of dictionaries which declares the details of every relevant measure to extract. The template for this list can be printed with the `template()` function. A more in depth explanation is provided further down this file.
9. `markColumn`, the column in which the marks are written in the individual .xlsx files. This is only known _after_ converting at least one file, since the position of the column changes depending on the number of arrays which are used in MedPC.
10. `timeColumn`, the column in which the time is written in the individual .xlsx files. This is only known _after_ converting at least one file, since the position of the column changes depending on the number of arrays which are used in MedPC.
11. `relocate`, a boolean (that is, it takes only values of `True` and `False`) which indicates whether or not the raw MedPC files should be moved from the temporary directory to the permanent one after the analysis. This is useful so as to avoid having to manualy move the files back to the temporary directory while the code is being tested and debugged.

The `timeColumn` and `markColumn` arguments are not needed to initialize the `Analyzer` object. The values for these arguments are obtained after first initializing the object without them and using the `.convert()` method to convert at least one file to .xlsx format:

```python
analyzer = Analyzer(fileName=summary_file, temporaryDirectory=temporary_directory, permanentDirectory=raw_directory,
                    convertedDirectory=converted_directory, subjectList=subjects, suffix="_", sheets=sheets,
                    analysisList=analysis_list, relocate=False)

analyzer.convert()
```


Then, this file must be manually inspected in order to get the letters of the columns which contain both the marks and the time registry. These columns are next to each other and are the same lenght, and are likely to be the longest columns in the entire file. 

![get_columns](https://user-images.githubusercontent.com/87039101/154622118-d96b7011-21d8-4414-87b0-9b2fa7c5df6f.png)



After the column letters are obtained the `timeColumn` and `markColumn` arguments can be provided and the `Analyzer` object is now ready to extract data.

```python
analyzer = Analyzer(fileName=summary_file, temporaryDirectory=temporary_directory, permanentDirectory=raw_directory,
                    convertedDirectory=converted_directory, subjectList=subjects, suffix="_", sheets=sheets,
                    analysisList=analysis_list, timeColumn="O", markColumn="P", relocate=False)
```

Thus, an example of a full script ready to analyze with a single click could be as follows:

```python
from medpcpy import *

summary_file = 'Response_distribution.xlsx'
temporary_directory = '/home/usuario/Documents/Proyecto/Temporal/'
raw_directory = '/home/usuario/Documents/Proyecto/Brutos/'
converted_directory = '/home/usuario/Documents/Proyecto/Convertidos/'
sheets = ["Trials", "Responses", "Latencies", "Nosepokes",]
subjects = ["A1", "A2", "A3"]
trials_columns = {"A1": 2, "A2": 4, "A3": 6}
levers_columns = {"A1": 2, "A2": 5, "A3": 8}
latencies_columns = {"A1": 2, "A2": 7, "A3": 12}
nosepokes_columns = {"A1": 2, "A2": 6, "A3": 10}

analysis_list = [
	# Completed trials
    {"fetch": {"cell_row": 15,
               "cell_column": 2,
               "sheet": "Trials",
               "summary_column_dict": trials_columns,
               }},
	# Response distributions
    {"resp_dist": {"trial_start": 300, "trial_end": 300, "response": 200,
                   "bin_size": 1,
                   "bin_amount": 15,
		   "label": "Responses",
                   }},
	# Lever presses
    {"conteoresp": {"trial_start": 114, "trial_end": 180, "response": 202,
                    "header": "PalDiscRef",
                    "sheet": "Responses",
                    "column": 1,
                    "summary_column_dict": levers_columns,
		    "substract": True,
                    }},
	# Lever latencies
    {"conteolat": {"trial_start": 112, "response": 113,
                   "header": "LatPalDisc",
                   "sheet": "Latencies",
                   "column": 2,
                   "summary_column_dict": latencies_columns,
		   "statistic": "mean",
                   }},
	# Nosepoke responses
    {"conteototal": {"response": 301,
                     "header": "EscForzDiscRef",
                     "sheet": "Nosepokes",
                     "column": 3,
                     "summary_column_dict": nosepokes_columns,
                     }},
]

analyzer = Analyzer(fileName=summary_file, temporaryDirectory=temporary_directory, raw_directory=directorioBrutos,
                    convertedDirectory=converted_directory, subjectList=subjects, suffix="_", sheets=sheets,
                    analysisList=analysis_list, timeColumn="O", markColumn="P", relocate=False)

analyzer.complete_analysis()

```


## Librería _oop_funciones.py_

Esta librería provee una manera simple de analizar los datos entregados por el programa de [Med PC-IV](https://www.med-associates.com/). Contiene funciones útiles para tareas y análisis básicos. La librería solamente ha sido probada con los archivos entregados por Med PC-IV, por lo que desconozco su funcionamiento con versiones distintas. Agradeceré cualquier realimentación al respecto.

De manera general la librería escanea una carpeta en la que se encuentran los archivos de texto sin formato entregados por MedPC que se desean analizar. Con base en ellos determina los sujetos y sesiones por analizar, convierte los archivos a formato ".xlsx" y separa las listas crudas de datos en columnas más legibles, realiza los conteos de respuestas, latencias, o distribuciones de respuesta que el usuario declare, y finalmente escribe los resultados en archivos individuales para cada sujeto y en un archivo de resumen. Tras declarar todas las variables pertinentes, el análisis completo de uno o más días de sesiones experimentales y de uno o más sujetos puede realizarse con un clic.

Sin embargo, la librería requiere de la declaración de variables específicas y su llamada en forma de argumentos en las funciones pertinentes.

Las variables necesarias para utilizar la librería son:

* Una variable que contenga el nombre del archivo de resumen en forma de _string_ en que se guardarán los datos. El _string_ debe contener la extensión ".xlsx". Ejemplo:
```python
archivo_de_resumen = "Igualacion.xlsx"
```

* Una variable que contenga la [dirección absoluta](https://www.geeksforgeeks.org/absolute-relative-pathnames-unix/) del directorio temporal en que se almacenarán los datos brutos antes de su análisis. Es importante que el último caracter del _string_ sea una diagonal `/`, y que cada nivel de la dirección sea separado por diagonales hacia adelante "`/`" y no hacia atrás "`\`". Ejemplo:
```python
directorioTemporal = "C:/Users/Admin/Desktop/Direccion/De/Tu/Carpeta/DirectorioTemporal/"  # En el caso de Windows
directorioTemporal = "/home/usuario/Documents/Direccion/De/Tu/Carpeta/DirectorioTemporal/"  # En el caso de Unix
```

* Una variable que contenga la dirección absoluta del directorio permanente en que se guardarán los datos brutos después de haber sido utilizados (no es necesario mover los archivos manualmente después de su utilización; el programa se encarga de eso automáticamente). Ejemplo:
```python
directorioBrutos = "C:/Users/Admin/Desktop/Direccion/De/Tu/Carpeta/DirectorioBrutos/"  # En el caso de Windows
directorioBrutos = "/home/usuario/Documents/Direccion/De/Tu/Carpeta/DirectorioBrutos/"  # En el caso de Unix
```

* Una variable que contenga la dirección absoluta del directorio en que se guardarán los datos convertidos después del análisis. En este directorio se almacenarán tanto los archivos individuales con extensión ".xlsx" como el archivo de resumen. Ejemplo:
```python
directorioConvertidos = "C:/Users/Admin/Desktop/Direccion/De/Tu/Carpeta/DirectorioConvertidos/"  # En el caso de Windows
directorioConvertidos = "/home/usuario/Documents/Direccion/De/Tu/Carpeta/DirectorioConvertidos/"  # En el caso de Unix
```

* Una lista de _strings_ con los nombres de las hojas de cálculo que debe contener el archivo de resumen. Ejemplo:
```python
hojas = ["RespuestasPalanca", "LatenciasPalanca", "RespuestasNosepoke"]
```

* Una lista con los nombres de los sujetos en forma de _string_. Ejemplo:
```python
sujetos = ["Rata1", "Rata2", "Rata3", "Rata4"]
```

* Uno o más [diccionarios](https://www.w3schools.com/python/python_dictionaries.asp) que relacionen a cada sujeto con la columna en que sus datos se escribirán en cada hoja del archivo de resumen. Es decir: un mismo sujeto puede tener asociadas múltiples medidas (e.g., respuestas en palancas, respuestas en nosepoke, latencias, etc). Además, medidas distintas pueden tener subdivisiones diferentes (e.g., puede haber dos medidas de respuestas a palancas (izquierda y derecha), y una sola medida para respuestas a nosepoke). Así, es posible que en una hoja se encuentre un formato similar a este:

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;![image](https://user-images.githubusercontent.com/87039101/140565456-64e9654d-c711-45dd-962f-f6e91b3af9a5.png)


&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Mientras que en otra se puede encontrar un formato similar a este: 

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;![image](https://user-images.githubusercontent.com/87039101/140565654-eb234a07-bb0b-464e-ae99-32faf808d86c.png)


&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Como se puede ver, medidas distintas para un mismo sujeto requerirían de cantidades distintas de columnas. Por ello será necesario declarar al menos dos diccionarios: uno que relacione a los sujetos con el espacio que ocupan en la primera hoja (respuestas en palancas), y otro que los relacione con el espacio que ocupan en la segunda hoja (respuestas en nosepoke). Estos diccionarios solamente necesitan declarar la primera columna ocupada traduciendo su letra en número (A = 1, B = 2, etc). El resto de las columnas es manejado más adelante. Así, un ejemplo de diccionarios sería:
```python
columnas_palancas = {"Rata1": 2, "Rata2": 7, "Rata3": 12,}
columnas_nosepoke = {"Rata1": 2, "Rata2": 5, "Rata3": 8,}
```

* Finalmente, el corazón de la librería es una lista de diccionarios que dictará las medidas que serán extraídas de los marcadores y del tiempo, al igual que la manera de escribirlas en los archivos individuales y de resumen. Esta lista tiene un formato específico que puede obtenerse ejecutando la función `template()` incluida con la librería. Cada diccionario de la lista declara la función a utilizar (copiar directamente de una celda, contar respuestas por ensayo, contar respuestas totales, contar latencias, contar respuestas por _bin_ de tiempo), junto con sus parámetros pertinentes (marcadores, columna en que se escribirán los datos en los archivos individuales y de resumen, etc). Ejemplo:
```python
analysis_list = [
    {"conteoresp": {"inicio_ensayo": 111, "fin_ensayo": 222, "respuesta": 333,
                    "column": 1,
                    "header": "Palanca Izq",
                    "sheet": "Palancas",
                    "summary_column_list": columnas_palancas,
		    "substract": True,
                    }},
    {"conteoresp": {"inicio_ensayo": 444, "fin_ensayo": 555, "respuesta": 666,
                    "column": 2,
                    "header": "Palanca Der",
                    "sheet": "Palancas",
                    "summary_column_list": columnas_palancas,
                    "substract": True,
		    "offset": 1,
                    }},
    {"conteototal": {"respuesta": 777,
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

Esta función permite extraer directamente un único dato de los archivos ".xlsx" individuales. Es funcional, por ejemplo, para extraer rápidamente el número de ensayos completados (si es que éste se encuentra dentro de alguna de las listas otorgadas por MedPC). Los argumentos `"cell_row"` y `"cell_column"` dictan la posición de la celda que se quiere extraer: un dato que se encuentra, por ejemplo, en la celda "B15" requerirá de los argumentos `"cell_row": 15` y `"cell_column": 2`. 

![image](https://user-images.githubusercontent.com/87039101/140596672-7213c34d-061c-4d05-a4bc-eced2abb65c5.png)

Los argumentos `"sheet"` y `"summary_column_list"` determinan la manera en que el dato extraído se escribirá en el archivo de resumen. El argumento `"sheet"` señala el nombre de la hoja de cálculo en que se escribirá el dato. Este nombre debe corresponder con uno de los elementos de la lista de hojas de cálculo generada anteriormente. Mientras, `"summary_column_list"` será el diccionario que asocia a sujetos con columnas (explicado anteriormente).

El argumento `"offset"` es un caso especial: en ocasiones se requerirá que medidas similares de un mismo sujeto sean escritas en columnas adyacentes de una misma hoja. Por ejemplo:

![image](https://user-images.githubusercontent.com/87039101/140452655-3109ae8b-3256-4596-93fe-d0b78196fd59.png)

En casos así no es necesario declarar diccionarios adicionales que asocien cada medida con una columna específica. Una manera más económica es declarar en varias medidas un único diccionario que asocie al sujeto con una columna "base", y después agregar incrementalmente unidades al argumento `"offset"`. Cada unidad en `"offset"` desplazará la medida en cuestión una columna hacia la derecha. Por ejemplo:
```python
analysis_list = [
{"fetch": {"cell_row": 10,
           "cell_column": 10,
           "sheet": "Sheet_1",
           "summary_column_list": column_dictionary,
           "offset": 0  # Innecesario
           }},
{"fetch": {"cell_row": 20,
           "cell_column": 20,
           "sheet": "Sheet_1",
           "summary_column_list": column_dictionary,
           "offset": 1  # <------
           }},
{"fetch": {"cell_row": 30,
           "cell_column": 30,
           "sheet": "Sheet_1",
           "summary_column_list": column_dictionary,
           "offset": 2  # <------
           }},
]
```

Esto resultaría en tres columnas: la primera se encontraría en la posición declarada por el diccionario `column_dictionary`, mientras que las otras dos se encontrarían una y dos posiciones a la derecha. 

Finalmente, si dentro del diccionario no se declara ningún valor para `"offset"`, éste tomará un valor por defecto de 0.

### Conteoresp

```python
analysis_list = [
    {"conteoresp": {"measures": 2, # Opcional
                    "inicio_ensayo": 111, "fin_ensayo": 222, "respuesta": 333,
                    "inicio_ensayo2": 444, "fin_ensayo2": 555, "respuesta2": 666, # Opcional
                    "column": 1,
                    "header": "Generic_title",
                    "sheet": "Sheet_2",
                    "summary_column_list": column_dictionary2,
		    "substract": True, # Opcional
                    "statistic": "mean",  # Opcional. Alternativa: "median"
		    "offset": 0,  # Opcional
                    }},
]
```

Esta función permite contar la cantidad de respuestas que ocurren entre el inicio y el fin de un tipo de ensayo particular. Una lista con todas las respuestas por ensayo se escribe en el archivo individual ".xlsx", y una medida de tendencia central (media o mediana) por sesión se escribe en el archivo de resumen.

Los argumentos `"inicio_ensayo"`, `"fin_ensayo"`, y `"respuesta"` son los marcadores de inicio de ensayo, fin de ensayo, y respuesta de interés, respectivamente.

El argumento `"substract"` es un argumento opcional que contempla el caso en el cual la respuesta que interesa contar sea también la respuesta que le da inicio al ensayo. En tal situación contar esa respuesta adicional resultaría en una imprecisión consistente que sobrestimaría en una unidad la cantidad total de respuestas por ensayo. Para solucionar esa situación, añadir el argumento `"substract": True` resultará en la resta de una unidad a cada uno de los conteos de respuestas por ensayo, lo cual nos devolverá a un conteo exacto. En caso de no ser necesario, no delcarar el argumento lo hace tomar un valor por defecto de `False`, con lo que no se realizará la resta.

Los argumentos `"column"` y `"header"` determinan la manera en que la lista completa de respuestas por ensayo se escribirá en el archivo individual ".xlsx". `"column"` indica la columna en la cual se pegará la lista (siendo que 1 = A, 2 = B, 3 = C, etc). El argumento `"header"` indica el rótulo que tendrá esa columna en su primera celda. 

Los argumentos `"sheet"` y `"summary_column_list"` determinan la manera en que la media de respuestas por ensayo de la sesión se escribirá en el archivo de resumen. El argumento `"sheet"` señala el nombre de la hoja de cálculo en que se escribirá el dato. Mientras, `"summary_column_list"` será el diccionario que asocia a sujetos con columnas (explicado anteriormente).

Esta función, junto con las funciones `conteolat` y `conteototal`, incorpora la posibilidad de realizar medidas "agregadas" o múltiples mediante el argumento `"measures"`: en algunas ocasiones es ventajoso sumar en una sola medida las respuesas (o latencias) provenientes de dos fuentes distintas. Como ejemplo se puede pensar en un caso en el cual haya respuestas en una palanca que lleven probabilísticamente a dos consecuencias diferentes y que, por descuido o planeación, tengan marcadores distintos. Las respuestas en ese caso deberán sumarse y contribuir a la misma media en el archivo de resumen. Para casos como ese el argumento `"measures"` permite agregar dentro de una misma medida fuentes distintas de información. `"measures"` indicará la cantidad de fuentes que se deberán agregar en la misma medida. Para cada medida adicional se deberán declarar además los marcadores pertinentes siguiendo la numeración lógica. Por ejemplo, para tres fuentes agregadas en una misma medida los argumentos serían:
```python
analysis_list = [
    {"conteoresp": {"measures": 3,
                    "inicio_ensayo": 123, "fin_ensayo": 124, "respuesta": 125,
                    "inicio_ensayo2": 223, "fin_ensayo2": 224, "respuesta2": 225, 
                    "inicio_ensayo3": 323, "fin_ensayo3": 324, "respuesta3": 325, 
                    ...
```

La atención debe centrarse en los dígitos 2 y 3 que siguen a los argumentos, notando que la numeración es consecutiva y que para la primera medida no se debe declarar el dígito 1. Esta función puede manejar una cantidad indefinida de fuentes aglomeradas en una misma medida.

Finalmente, el argumento `"statistic"` determina la medida de tendencia central (media o mediana) que será escrita en el archivo de resumen. Su valor por defecto es `"mean"`, por lo que si no se declara ningún valor, la medida escrita será la media.

### Conteototal

```python
 analysis_list = [
    {"conteototal": {"measures": 2, # Opcional
                     "respuesta": 111,
                     "respuesta2": 222, # Opcional
                     "column": 3,
                     "header": "Generic_title",
                     "sheet": "Sheet_4",
                     "summary_column_list": column_dictionary4,
                     "offset": 0,  # Opcional
                     }},
]
```

Esta función permite contar la cantidad total de respuestas en toda una sesión sin diferenciar entre ensayos. El número de ocurrencias de la respuesta se escribe tanto en el archivo individual ".xlsx" como en el archivo de resumen.

Sus argumentos son idénticos a los de `conteoresp()` salvo por dos excepciones: tiene un único argumento para marcador (el marcador de la respuesta de interés), por lo que al agregar más de una fuente en la misma medida mediante `"measures"` la numeración saltará de uno en uno; y no tiene el argumento de `"substract"` en tanto que no hay una respuesta extra por descartar.

### Conteolat

```python
analysis_list = [
    {"conteolat": {"measures": 2, # Opcional
                   "inicio_ensayo": 111, "respuesta": 222,
                   "inicio_ensayo2": 333, "respuesta2": 444, # Opcional
                   "column": 2,
                   "header": "Generic_title",
                   "sheet": "Sheet_3",
                   "summary_column_list": column_dictionary3,
		   "statistic": "mean",  # Opcional. Alternativa: "median"
                   "offset": 0,  # Opcional
		   "unit": 20
                   }},
]
```

Esta función permite contar las latencias por ensayo medidas en segundos desde el inicio del ensayo hasta la primera ocurrencia de la respuesta de interés. La lista completa con las latencias de respuesta de cada ensayo se escribe en el archivo individual ".xlsx", y el estadístico elegido (media o mediana) se escribe en el archivo de resumen.

Al igual que `conteoresp`, esta función incorpora el argumento `"statistic"` para determinar la medida de tendencia central escrita en el archivo de resumen.

Esta función incorpora el argumento `unit`, que determina la resolución temporal utilizada en Med. El argumento corresponde con las unidades entre las que se divide cada segundo. Si, por ejemplo, la resolución temporal utilizada es de vigésimas de segundo entonces el argumento `unit` deberá tomar el valor de 20.

### Resp_dist

```python
analysis_list = [
    {"resp_dist": {"inicio_ensayo": 111, "fin_ensayo": 222, "respuesta": 333,
                   "bin_size": 1,
                   "bin_amount": 15,
		   "label": "Generic_title",  # Opcional
		   "statistic": "mean",  # Opcional. Alternativa: "median"
		   "unit": 20,
                   }},
]
```

Esta función permite determinar la distribución temporal de una respuesta de interés a lo largo de cada uno de los ensayos de una sesión. El programa dividirá cada ensayo en _bins_, después contará la cantidad de ocasiones que la respuesta de interés ocurrió en cada uno de los _bins_ y almacenará la información en listas. Cada ensayo generará una lista separada, y todas las listas serán escritas en una misma hoja del archivo individual ".xlsx" distinta de aquella en que se escribe el resto de las listas generadas por las otras funciones. Además, una lista con las medias de respuestas por bin se escribirá en una columna en una hoja del archivo de resumen. Cada sujeto tendrá una hoja exclusiva que será generada automáticamente por el programa y cada sesión ocupará una columna en esa hoja. En tanto que cada distribución de respuestas para cada sujeto es escrita en una hoja separada generada automáticamente, esta función no necesita de los argumentos `"column"`, `"header"`, `"sheet"`, y `"summary_column_list"`.

En aquellos casos en que no haya intervalo entre ensayos y no exista un marcador de fin de ensayo, sino que el fin de un ensayo sea señalado solamente por el inicio del ensayo siguiente, bastará con declarar el mismo marcador para los argumentos `"inicio_ensayo"` y `"fin_ensayo"`.

Los argumentos `"bin_size"` y `"bin_amount"` determinan la duración en segundos de cada _bin_ y la cantidad de _bins_ por ensayo, respectivamente. Así, un ensayo de 15 segundos con _bins_ de un segundo tendrá como argumentos `"bin_size": 1` y `"bin_amount": 15`.

El programa crea un _bin_ adicional a los declarados con `"bin_amount"` en el cual se aglomeran todas las respuestas que ocurran más allá del fin del último _bin_ declarado. De no haber tales respuestas el _bin_ final resultará vacío.

En caso de que se requiera obtener la distribución de más de una respuesta del mismo experimento se debe declarar el argumento opcional `"label"` con un nombre que identifique a cada una de las medidas que se requieren. El programa creará una hoja separada para cada medida de cada sujeto y le pondrá como título el nombre del sujeto seguido del _string_ que se haya utilizado como argumento de `"label"`, y cada sesión ocupará una columna en su hoja pertinente. Por ejemplo, si se requiere obtener distribuciones de respuestas en palancas y en nosepoke, los diccionarios necesarios podrían tener un formato como este:

```python 
analysis_list = [
    {"resp_dist": {"inicio_ensayo": 111, "fin_ensayo": 222, "respuesta": 333,
                   "bin_size": 1,
                   "bin_amount": 15,
		   "label": "Palancas", # Opcional
		   "statistic": "mean",
		   "unit": 20,
                   }},

    {"resp_dist": {"inicio_ensayo": 444, "fin_ensayo": 555, "respuesta": 666,
                   "bin_size": 1,
                   "bin_amount": 15,
		   "label": "Nosepoke",  # Opcional
		   "statistic": "mean",
		   "unit": 20,
                   }},
]
```

Y el archivo de resumen resultante tendría dos hojas para cada sujeto: una asignada a las distribuciones de respuestas en palancas y otra asignada a las distribuciones de respuesta en nosepoke. Si los sujetos fuesen `"Rata1"` y `"Rata2"`, las hojas resultantes tendrían los nombres de `"Rata1_Palancas"`, `"Rata1_Nosepoke"`, `"Rata2_Palancas"`, y `"Rata2_Nosepoke"`. Por otro lado, el archivo ".xlsx" individual de cada sujeto contendría dos hojas: una para cada medida. Estas hojas son creadas automáticamente y llevan por título el valor del argumento `"label"`.

En caso de que el argumento `"label"` sea omitido se creará en el archivo de resumen una sola hoja por sujeto, y ésta llevará por título el nombre del sujeto. Si se declaran múltiples distribuciones de respuestas y en todas se omite el argumento `"label"`, éstas se sobreescribirán entre sí y solamente será visible la última medida declarada.

Esta función, al igual que `conteoresp` y `conteolat`, incorpora el argumento `"statistic"` para determinar la medida de tendencia central escrita en el archivo de resumen.

Al igual que `conteolat` esta función incorpora el argumento `unit` para dictar la resolución temporal del análisis.


___
___
## Uso

Como primer paso será necesario importar la librería al proyecto actual e instalar las dependencias pertinentes, es decir, Openpyxl y Pandas.

La importación de la librería puede realizarse descargando el archivo oop_funciones.py de este github y guardándolo en la misma carpeta en que se encuentre el script de python que se esté construyendo. Después, como primera línea del script se debe escribir
```python
from oop_funciones import *
```

para importar todas las funciones de la librería.

Tras la declaración de todas las variables pertinentes será necesario solamente crear un objeto de tipo `Analyzer` y asignarlo a una variable con los argumentos adecuados. Los argumentos necesarios son:

1. `fileName`, el nombre del archivo de resumen.
2. `temporaryDirectory`, el directorio temporal declarado antes en el cual se almacenan los datos antes de su análisis.
3. `permanentDirectory`, el directorio al que se moverán los datos brutos después de su utilización.
4. `convertedDirectory`, el directorio en el que se guardarán los archivos individuales convertidos ".xlsx" y el archivo de resumen.
5. `subjectList`, la lista con los nombres de los sujetos.
6. `suffix`, un string que indica al programa el caracter o conjunto de caracteres que separa el nombre de los sujetos del número de la sesión en el nombre de los archivos. Se recomienda un valor de `"_"`.
7. `sheets`, una lista de _strings_ con los nombres de las hojas de cálculo que deberán crearse dentro del archivo de resumen.
8. `analysisList`, la lista de diccionarios que contiene los análisis a realizar.
9. `markColumn`, la columna de los archivos individuales en la cual se escribieron los marcadores.
10. `timeColumn`, la columna de los archivos individuales en la cual se escribió el tiempo asociado a cada marcador.
11. `relocate`, un booleano (es decir, toma valores de `True` y `False`) que indica si los archivos crudos deberán ser movidos del directorio temporal al directorio permanente después del análisis. Es útil para evitar la necesidad de regresar los archivos manualmente al directorio temporal mientras se están haciendo pruebas con el código.

Los argumentos `timeColumn` y `markColumn` no son necesarios inicialmente, sino que son obtenidos al aplicar parte del programa a los datos. Esto se verá a continuación.

La creación inicial del objeto de tipo `Analyzer` puede ser como sigue (apelando a las variables creadas anteriormente):
```python
analyzer = Analyzer(fileName=archivo_de_resumen, temporaryDirectory=directorioTemporal, permanentDirectory=directorioBrutos,
                    convertedDirectory=directorioConvertidos, subjectList=sujetos, suffix="_", sheets=hojas,
                    analysisList=analysis_list, relocate=False)
```

Sin embargo, aun no sería posible hacer el análisis completo de los datos, sino solamente su conversión, que es justamente el paso siguiente. Con el método `convert()` se pueden convertir los archivos contenidos en la carpeta temporal:

```python
analyzer.convert()
```

Esto generará los archivos individuales en formato .xlsx. Será necesario ahora abrir cualquiera de ellos con cualquier editor de hojas de cálculo y determinar manualmente las columnas en las cuales se escribieron las listas de tiempo y marcadores. Estas columnas serán después pasadas como argumentos en la declaración del objeto tipo `Analyzer`. Por ejemplo, suponiendo que en los archivos individuales encontrásemos que la lista con el tiempo fue escrita en la columna "M" y la lista con los marcadores en la columna "N", la declaración del objeto resultaría finalmente como:

```python
analyzer = Analyzer(fileName=archivo, temporaryDirectory=directorioTemporal, permanentDirectory=directorioBrutos,
                    convertedDirectory=directorioConvertidos, subjectList=sujetos, suffix="_", sheets=hojas,
                    analysisList=analysis_list, timeColumn="M", markColumn="N", relocate=False)
```

Así, el script completo preparado para analizar datos con un solo clic sería:

```python
from oop_funciones import *

archivo = 'Response_distribution.xlsx'
directorioTemporal = '/home/usuario/Documents/Proyecto/Temporal/'
directorioBrutos = '/home/usuario/Documents/Proyecto/Brutos/'
directorioConvertidos = '/home/usuario/Documents/Proyecto/Convertidos/'
hojas = ["Ensayos", "Respuestas", "Latencias", "Nosepokes",]
sujetos = ["A1", "A2", "A3"]
columnasEnsayos = {"A1": 2, "A2": 4, "A3": 6}
columnasRespuestas = {"A1": 2, "A2": 5, "A3": 8}
columnasLatencias = {"A1": 2, "A2": 7, "A3": 12}
columnasNosepokes = {"A1": 2, "A2": 6, "A3": 10}

analysis_list = [
	# Ensayos completados
    {"fetch": {"cell_row": 15,
               "cell_column": 2,
               "sheet": "Ensayos",
               "summary_column_list": columnasEnsayos,
               }},
	# Distribucion respuestas
    {"resp_dist": {"inicio_ensayo": 300, "fin_ensayo": 300, "respuesta": 200,
                   "bin_size": 1,
                   "bin_amount": 15,
		   "label": "Respuestas",
                   }},
	# Respuestas palancas
    {"conteoresp": {"inicio_ensayo": 114, "fin_ensayo": 180, "respuesta": 202,
                    "header": "PalDiscRef",
                    "sheet": "Respuestas",
                    "column": 1,
                    "summary_column_list": columnasRespuestas,
		    "substract": True,
                    }},
	# Latencias palancas
    {"conteolat": {"inicio_ensayo": 112, "respuesta": 113,
                   "header": "LatPalDisc",
                   "sheet": "Latencias",
                   "column": 2,
                   "summary_column_list": columnasLatencias,
		   "statistic": "mean",
                   }},
	# Respuestas nosepokes
    {"conteototal": {"respuesta": 301,
                     "header": "EscForzDiscRef",
                     "sheet": "Nosepokes",
                     "column": 3,
                     "summary_column_list": columnasNosepokes,
                     }},
]

analyzer = Analyzer(fileName=archivo, temporaryDirectory=directorioTemporal, permanentDirectory=directorioBrutos,
                    convertedDirectory=directorioConvertidos, subjectList=sujetos, suffix="_", sheets=hojas,
                    analysisList=analysis_list, timeColumn="M", markColumn="N", relocate=False)

analyzer.complete_analysis()

```
