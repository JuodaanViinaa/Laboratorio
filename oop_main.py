from oop_funciones import *

archivo = 'Prueba_mean.xlsx'
directorioTemporal = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Temporal/'
directorioBrutos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/PruebasBrutos/'
directorioConvertidos = '/home/daniel/Documents/Doctorado/Proyecto de Doctorado/ExperimentoEscape/Flex/'

sujetos = ['E3', 'E4', 'E5', 'E7', 'E8', 'E9', 'E1']
columnasProp = {"E3": 2, "E4": 4, "E5": 6, "E7": 8, "E8": 10, "E9": 12, "E1": 14}
columnasResp = {"E3": 2, "E4": 9, "E5": 16, "E7": 23, "E8": 30, "E9": 37, "E1": 44}
columnasLatPal = {"E3": 2, "E4": 7, "E5": 12, "E7": 17, "E8": 22, "E9": 27, "E1": 32}
columnasEscapes = {"E3": 2, "E4": 13, "E5": 24, "E7": 35, "E8": 46, "E9": 57, "E1": 68}
columnasLatEsc = {"E3": 2, "E4": 13, "E5": 24, "E7": 35, "E8": 46, "E9": 57, "E1": 68}
columnasEscForz = {"E3": 2, "E4": 7, "E5": 12, "E7": 17, "E8": 22, "E9": 27, "E1": 32}
hojas = ['Proporciones', 'Respuestas', 'Latencias', 'Comedero', 'Escapes', 'LatNosepoke', 'EscapesForzados',
         'LatEscapeForz', 'LatEscapeForzPorEstim']


analysis_list = [
    # Proporciones
    {"fetch": {"cell_row": 14,
               "cell_column": 6,
               "sheet": "Proporciones",
               "summary_column_list": columnasProp,
               }},
    # Ensayos completados
    {"fetch": {"cell_row": 15,
               "cell_column": 2,
               "sheet": "Proporciones",
               "summary_column_list": columnasProp,
               "offset": 1
               }},

    # Respuestas palancas
    # Respuestas Pal forzados discriminativos reforzados
    {"conteoresp": {"inicio_ensayo": 114, "fin_ensayo": 180, "respuesta": 202,
                    "header": "PalDiscRef",
                    "sheet": "Respuestas",
                    "column": 1,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    }},
    # Respuestas Pal forzados discriminativos no reforzados
    {"conteoresp": {"inicio_ensayo": 115, "fin_ensayo": 180, "respuesta": 202,
                    "header": "PalNoDiscRef",
                    "sheet": "Respuestas",
                    "column": 3,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 1,
                    }},
    # Respuestas Pal forzados no discriminativos 1
    {"conteoresp": {"inicio_ensayo": 134, "fin_ensayo": 180, "respuesta": 201,
                    "header": "PalNoDisc1",
                    "sheet": "Respuestas",
                    "column": 5,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 2,
                    }},
    # Respuestas Pal forzados no discriminativos 2
    {"conteoresp": {"inicio_ensayo": 137, "fin_ensayo": 180, "respuesta": 201,
                    "header": "PalNoDisc2",
                    "sheet": "Respuestas",
                    "column": 7,
                    "substract": True,
                    "summary_column_list": columnasResp,
                    "offset": 3,
                    }},

    # Latencias palancas
    # Latencias pal forzados discriminativos
    {"conteolat": {"inicio_ensayo": 112, "respuesta": 113,
                   "header": "LatPalDisc",
                   "sheet": "Latencias",
                   "column": 9,
                   "summary_column_list": columnasLatPal,
                   "statistic": "mean",
                   }},
    # Latencias pal forzados no discriminativos
    {"conteolat": {"inicio_ensayo": 132, "respuesta": 133,
                   "header": "LatPalNoDisc",
                   "sheet": "Latencias",
                   "column": 11,
                   "summary_column_list": columnasLatPal,
                   "offset": 1,
                   "statistic": "mean",
                   }},

    # Respuestas comederos
    # Respuestas comederos forzados discriminativos reforzados
    {"conteoresp": {"inicio_ensayo": 114, "fin_ensayo": 16, "respuesta": 203,
                    "header": "ComDiscRef",
                    "sheet": "Comedero",
                    "column": 13,
                    "summary_column_list": columnasResp,
                    }},
    # Respuestas comederos forzados discriminativos no reforzados
    {"conteoresp": {"inicio_ensayo": 115, "fin_ensayo": 117, "respuesta": 203,
                    "header": "ComNoDiscRef",
                    "sheet": "Comedero",
                    "column": 15,
                    "summary_column_list": columnasResp,
                    "offset": 1,
                    }},
    # Respuestas comederos forzados no discriminativos 1
    {"conteoresp": {"inicio_ensayo": 134, "fin_ensayo": 40, "respuesta": 203,
                    "header": "ComNoDisc1",
                    "sheet": "Comedero",
                    "column": 17,
                    "summary_column_list": columnasResp,
                    "offset": 2,
                    }},
    # Respuestas comederos forzados no discriminativos 2
    {"conteoresp": {"inicio_ensayo": 137, "fin_ensayo": 43, "respuesta": 203,
                    "header": "ComNoDisc2",
                    "sheet": "Comedero",
                    "column": 19,
                    "summary_column_list": columnasResp,
                    "offset": 3,
                    }},

    # Escapes
    # Respuestas nosepoke forzados discriminativos reforzados
    {"conteototal": {"respuesta": 301,
                     "header": "EscForzDiscRef",
                     "sheet": "Escapes",
                     "column": 21,
                     "summary_column_list": columnasEscapes,
                     }},
    # Respuestas nosepoke forzados discriminativos no reforzados
    {"conteototal": {"respuesta": 302,
                     "header": "EscForzDiscNoRef",
                     "sheet": "Escapes",
                     "column": 23,
                     "summary_column_list": columnasEscapes,
                     "offset": 1,
                     }},
    # Respuestas nosepoke forzados no discriminativos 1
    {"conteototal": {"respuesta": 303,
                     "header": "EscForzNoDisc1",
                     "sheet": "Escapes",
                     "column": 25,
                     "summary_column_list": columnasEscapes,
                     "offset": 2,
                     }},
    # Respuestas nosepoke forzados no discriminativos 2
    {"conteototal": {"respuesta": 304,
                     "header": "EscForzNoDisc2",
                     "sheet": "Escapes",
                     "column": 27,
                     "summary_column_list": columnasEscapes,
                     "offset": 3,
                     }},
    # Respuestas nosepoke forzados discriminativos reforzados
    {"conteototal": {"respuesta": 305,
                     "header": "EscLibDiscRef",
                     "sheet": "Escapes",
                     "column": 29,
                     "summary_column_list": columnasEscapes,
                     "offset": 4,
                     }},
    # Respuestas nosepoke forzados discriminativos no reforzados
    {"conteototal": {"respuesta": 306,
                     "header": "EscLibDiscNoRef",
                     "sheet": "Escapes",
                     "column": 31,
                     "summary_column_list": columnasEscapes,
                     "offset": 5,
                     }},
    # Respuestas nosepoke forzados no discriminativos 1
    {"conteototal": {"respuesta": 307,
                     "header": "EscLibNoDisc1",
                     "sheet": "Escapes",
                     "column": 33,
                     "summary_column_list": columnasEscapes,
                     "offset": 6,
                     }},
    # Respuestas nosepoke forzados no discriminativos 2
    {"conteototal": {"respuesta": 308,
                     "header": "EscLibNoDisc2",
                     "sheet": "Escapes",
                     "column": 35,
                     "summary_column_list": columnasEscapes,
                     "offset": 7,
                     }},

    # Latencias Escape
    # Latencias nosepoke forzados discriminativos reforzados
    {"conteolat": {"inicio_ensayo": 114, "respuesta": 301,
                   "header": "LatEscFDiscRef",
                   "sheet": "LatNosepoke",
                   "column": 37,
                   "summary_column_list": columnasEscapes,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke forzados discriminativos no reforzados
    {"conteolat": {"inicio_ensayo": 115, "respuesta": 302,
                   "header": "LatEscFDiscNoRef",
                   "sheet": "LatNosepoke",
                   "column": 39,
                   "summary_column_list": columnasEscapes,
                   "offset": 1,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke forzados no discriminativos 1
    {"conteolat": {"inicio_ensayo": 134, "respuesta": 303,
                   "header": "LatEscFNoDisc1",
                   "sheet": "LatNosepoke",
                   "column": 41,
                   "summary_column_list": columnasEscapes,
                   "offset": 2,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke forzados no discriminativos 2
    {"conteolat": {"inicio_ensayo": 137, "respuesta": 304,
                   "header": "LatEscFNoDisc2",
                   "sheet": "LatNosepoke",
                   "column": 43,
                   "summary_column_list": columnasEscapes,
                   "offset": 3,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"inicio_ensayo": 154, "respuesta": 305,
                   "header": "LatEscLDiscRef",
                   "sheet": "LatNosepoke",
                   "column": 45,
                   "summary_column_list": columnasEscapes,
                   "offset": 4,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"inicio_ensayo": 155, "respuesta": 306,
                   "header": "LatEscLDiscRef",
                   "sheet": "LatNosepoke",
                   "column": 47,
                   "summary_column_list": columnasEscapes,
                   "offset": 5,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"inicio_ensayo": 157, "respuesta": 307,
                   "header": "LatEscLNoDisc1",
                   "sheet": "LatNosepoke",
                   "column": 49,
                   "summary_column_list": columnasEscapes,
                   "offset": 6,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke libres discriminativos reforzados
    {"conteolat": {"inicio_ensayo": 160, "respuesta": 308,
                   "header": "LatEscLDisc2",
                   "sheet": "LatNosepoke",
                   "column": 51,
                   "summary_column_list": columnasEscapes,
                   "offset": 7,
                   "statistic": "mean",
                   }},

    # Escapes forzados
    # Respuestas nosepoke escape forzado discriminativo
    {"conteototal": {"respuesta": 403,
                     "header": "EscForzDisc",
                     "sheet": "EscapesForzados",
                     "column": 53,
                     "summary_column_list": columnasEscForz,
                     }},
    # Respuestas nosepoke escape forzado no discriminativo
    {"conteototal": {"respuesta": 406,
                     "header": "EscForzNoDisc",
                     "sheet": "EscapesForzados",
                     "column": 55,
                     "summary_column_list": columnasEscForz,
                     "offset": 1,
                     }},

    # Latencias Escape Forzado
    # Latencias nosepoke escape forzado discriminativo
    {"conteolat": {"inicio_ensayo": 401, "respuesta": 403,
                   "header": "LatEscForzDisc",
                   "sheet": "LatEscapeForz",
                   "column": 57,
                   "summary_column_list": columnasLatPal,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke escape forzado no discriminativo
    {"conteolat": {"inicio_ensayo": 404, "respuesta": 406,
                   "header": "LatEscForzNoDisc",
                   "sheet": "LatEscapeForz",
                   "column": 59,
                   "summary_column_list": columnasLatPal,
                   "offset": 1,
                   "statistic": "mean",
                   }},

    # Latencias Escape Forzado por estímulo
    # Latencias nosepoke escape forzado discriminativo positivo
    {"conteolat": {"inicio_ensayo": 407, "respuesta": 403,
                   "header": "LatEscForzDiscPos",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 61,
                   "summary_column_list": columnasLatPal,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke escape forzado no discriminativo negativo
    {"conteolat": {"inicio_ensayo": 408, "respuesta": 403,
                   "header": "LatEscForzDiscNeg",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 63,
                   "summary_column_list": columnasLatPal,
                   "offset": 1,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke escape forzado discriminativo 1
    {"conteolat": {"inicio_ensayo": 409, "respuesta": 406,
                   "header": "LatEscForzNoDiscLuz1",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 65,
                   "summary_column_list": columnasLatPal,
                   "offset": 2,
                   "statistic": "mean",
                   }},
    # Latencias nosepoke escape forzado no discriminativo 2
    {"conteolat": {"inicio_ensayo": 410, "respuesta": 406,
                   "header": "LatEscForzNoDiscLuz2",
                   "sheet": "LatEscapeForzPorEstim",
                   "column": 67,
                   "summary_column_list": columnasLatPal,
                   "offset": 3,
                   "statistic": "mean",
                   }},

    # Análisis múltiples
    # Latencias Pal Disc
    {"conteolat": {"measures": 2,
                   "inicio_ensayo": 112, "respuesta": 113,  # Ensayos forzados discriminativos
                   "inicio_ensayo2": 401, "respuesta2": 402,  # Ensayos discriminativos de escape forzado
                   "header": "LatPalDiscTotal",
                   "sheet": "Latencias",
                   "column": 69,
                   "summary_column_list": columnasLatPal,
                   "statistic": "mean",
                   }},
    # Latencias Pal No Disc
    {"conteolat": {"measures": 2,
                   "inicio_ensayo": 132, "respuesta": 133,  # Ensayos forzados no discriminativos
                   "inicio_ensayo2": 404, "respuesta2": 405,  # Ensayos no discriminativos de escape forzado
                   "header": "LatPalNoDiscTotal",
                   "sheet": "Latencias",
                   "column": 71,
                   "summary_column_list": columnasLatPal,
                   "offset": 1,
                   "statistic": "mean",
                   }},
]

analyzer = Analyzer(fileName=archivo, temporaryDirectory=directorioTemporal, permanentDirectory=directorioBrutos,
                    convertedDirectory=directorioConvertidos, subjectList=sujetos, suffix="_SUBCHOIL_", sheets=hojas,
                    analysisList=analysis_list, timeColumn="O", markColumn="P", relocate=False)

analyzer.complete_analysis()
