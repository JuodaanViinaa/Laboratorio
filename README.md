# Laboratorio
Código utilizado en el laboratorio 101.

Estos _scripts_ se sirven de las librerías [Openpyxl](https://openpyxl.readthedocs.io/en/stable/index.html) y [Pandas](https://pandas.pydata.org/pandas-docs/stable/), por lo que será útil leer su documentación para entender algunas de las funciones utilizadas.

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
* `sesionInicial`, que indica el número de la primera sesión a convertir.
* `sesionFinal`, que indica el número de la última sesión a convertir.
* `sujetos`, que contendrá una lista con los nombres de los sujetos que forman el grupo a convertir.

Estas variables son las mismas utilizadas por el _script_ de resumen, de modo que solo es necesario declararlas una vez.

## ResumenUltimate.py

El _script_ está adaptado para el experimento de _escape_. 

Inicialmente, el código revisa una carpeta que contendrá temporalmente los archivos brutos a convertir y analizar. Busca los nombres de los sujetos declarados en la lista `sujetos` dentro de esta carpeta, y si algún sujeto no tiene ningún dato asociado, lo agrega a la lista `sujetosFaltantes`. Después se obtiene la sesión inicial y final de cada uno de los sujetos que sí tienen datos asociados con base en los nombres de los archivos. Se toma como sesión inicial al número de sesión más bajo, y como final al más alto. El código no puede manejar sesiones faltantes dentro de este rango (quizá en el futuro esto pueda resolverse utilizando una lista de listas en la cual cada sublista contenga todas las sesiones presentes para cada sujeto). 

Se compara la lista `sujetosFaltantes` con la lista `sujetos`. Todos aquellos sujetos que se encuentren en la lista `sujetosFaltantes` son eliminados junto con sus columnas asociadas en sus listas respectivas.

El siguiente paso es la conversión de los archivos a formato .xlsx. Después de convertir los archivos, el código los lee nuevamente y realiza un conteo de respuestas y latencias que escribe en dos lugares: en una hoja nueva del archivo individual generado por el convertidor (donde se incluyen respuestas y latencias por ensayo) y en un archivo de resumen con extensión .xlsx (donde solo se incluyen medias para las respuestas y medianas para las latencias).

El _script_ de resumen declara funciones para las tareas repetitivas:
* `hoja(nombre)`: crea hojas de cálculo en el archivo de resumen con el nombre dado como argumento. Si una hoja con ese nombre ya existe, no se crea una hoja nueva, sino que ésta simplemente se abre. Esta función debe ser asignada a una variable (e.g., `latencias = hoja('Latencias')`).
* `conteoresp(inicioEnsayo, finEnsayo, respuesta)`: cuenta respuestas por tipo de ensayo. Los argumentos son los marcadores de Med para inicio de ensayo, fin de ensayo, y respuesta de interés. Esta función resulta en una lista que contiene la cantidad de ocasiones que la respuesta ocurrió entre cada inicio y fin de ensayo. Si no ocurrieron ensayos del tipo de interés, la lista resultante tendrá un único cero para poder hacer cálculos de medias y medianas. Es importante tener en cuenta que generalmente los ensayos son iniciados por una respuesta. Esta primera respuesta ocurre después del marcador de inicio de ensayo, pero no forma parte del conteo de respuestas real. Dependiendo de que la respuesta que inicia el ensayo sea también la respuesta de interés a contar puede ser necesario restar una unidad al calcular la media de respuestas por ensayo. Esta función debe ser asignada a una variable.
* `conteototal(respuesta)`: cuenta respuestas totales de un tipo dado en la sesión completa. El argumento es el marcador de Med de la respuesta de interés. Esta función resulta en un número entero que representa la cantidad de ocasiones que la respuesta ocurrió en la sesión completa. Esta función debe ser asignada a una variable.
* `conteolat(inicioEnsayo, respuesta)`: cuenta la latencia entre el inicio de un ensayo de tipo determinado y la primera respuesta de interés. Los argumentos son los marcadores de Med de inicio de ensayo y respuesta. La función resulta en una lista con la latencia de respuesta en cada ensayo. Esta función debe ser asignada a una variable.
* `esccolumnas(titulo, columna, lista, restar)`: escribe las respuestas y latencias por ensayo en el archivo convertido individual. El argumento de `titulo` es un _string_ e indica el encabezado que tendrá la columna en donde se escribirá la lista; `columna` es un número entero e indica la columna en la cual se escribirá la lista (1 = A, 2 = B, etc.); `lista` indicará el nombre de la variable que contiene la lista que será escrita en la columna (las listas serán los resultados de las funciones descritas anteriormente); y `restar` indicará si se debe restar una unidad a cada elemento de la lista antes de pegarlo en su columna individual (debido a que la función `conteoresp()` cuenta una respuesta adicional por ensayo en situaciones ya descritas), y puede tomar los valores de `True` y `False`. Si el argumento `restar` es `True` y hay un ensayo con cero respuestas, no se ejecuta la resta para evitar resultar en números negativos.

La sección principal del _script_ se basa en dos ciclos `for`: uno que cicla a través de sesiones y otro anidado en el primero que lo hace a través de sujetos.

El _script_ genera listas con los conteos de respuestas y latencias de cada sujeto, y escribe medias y medianas en los sitios correspondietes del archivo de resumen. Los sitios correspondientes son determinados consultando las listas `columnasProp`, `columnasResp`, `columnasLatPal`, `columnasEscapes`, y `columnasLatEsc`. Existe una lista por cada una de las variables que se analizan debido a que cada variable se escribe en una hoja distinta y ocupa una cantidad de columnas distinta. Cada lista declara el número correspondiente a la primera columna utilizada por cada uno de los sujetos, por lo que habrá tantos elementos en cada lista como sujetos en la lista de `sujetos`. 

## Resumen(xls).py

Este resumen está desacoplado del convertidor, y está hecho para utilizarse con el convertidor anterior del laboratorio (que produce archivos convertidos con extensión .xls en lugar de .xlsx). Tiene las mismas funciones que _ResumenUltimate.py_ salvo porque no pega datos por ensayo en el archivo individual. Es preferible no usar este script.
