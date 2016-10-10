Autoxls
=======

`autoxls` es una herramienta de linea de comandos para automatizar las generación de archivos Excel 2003 a partir de resultados obtenidos de consultas a bases de datos.

# Contenidos

* [Características principales](#markdown-header-caracteriticas-principales)
* [Antes de empezar](#markdown-header-antes-de-empezar)
* [Primeros pasos](#markdown-header-primeros-pasos)
* [Ejecución](#markdown-header-ejecucion)
	* [Niveles de log](#markdown-header-niveles-de-log)
	* [Definiciones de keywords](#markdown-header-definiciones-de-keywords)
* [Notas para Desarrollo](#markdown-header-notas-para-desarrollo)

Características principales
===========================

* Múltiples conexiones a bases de datos
* Mútiples fuentes de datos
* Múltiples consultas/querys/Stored procedures
* Capacidad de procesar múltiples recordsets a partir de una única consulta
* Generáción automatizada de uno o más archivos Excel por ejecución
* Definición dinámica de nombres y textos a partir de "keywords" definidas en la invocación o por un archivo externo de keywords, puede aplicar:
	* Nombre del archivo
	* Nombres de las solapas
	* Textos en la planilla
	* Parámetros de invocación de la consulta
	* Datos de la conexión: servidor, usuario, contraseña
* Definir multiples solapas por archivo
* Automatizar la generación de múltiples objetos
	* textos y filas completas
	* formulas
	* grillas de datos o tablas
* Definir formatos de los objetos
	* anchos de columnas
	* colores
	* tipos de letra
	* Alineaciones
	* formatos númericos


Antes de empezar
================

Para comenzar, antes de probar la ejecución de esta herramienta, es necesario escribir un archivo de definición de la exportación a realizar. Este no es más que un archivo en formato JSON, que describe datos, archivos, planillas y formatos de la exportación. Nota: para validar el formato del mismo: [JSON Editor Online](http://www.jsoneditoronline.org/). 

Este archivo definirá los siguientes elementos:

* **datasources**: uno o más conexiones a bases de datos, con las siguientes propiedades:

	* **data\_connect\_str**: que establece una conexión a un servidor de datos.
	* **data\_query**: que define la correspondiente consulta. Podrá ser una "query" común o directamente la ejecución de un stored procedure.
	* **data\_query\_file**: Archivo dónde encontramos la consulta SQL.
	
* **files**:  Que define la generación de uno o más archivos Excel. Por cada archivo se podrá definir una o más:
	
	* **sheets**: Es decir, solapas de la planilla, por cada una de estas se pueden definir varios objetos Excel:
		
		* **text**: Texto estático, normalmente titulos, se define la celda, el texto y el formato.
		* **text_rows**: Define una lista de textos estáticos, que se escriben a partir de una posición en la fila, una celda a continuación de la otra con un determinado formato
		* **text_formated**: Una especialización de los objetos de texto, que permite aplicar formatos sobre datos recibidos en las __keywords__.
		* **datagrid**: Una grilla de datos, la salida final de los datos recuperados.
		* **formulas**: Formulas de escel
		* **table**: Una tabla Excel. 

* **formats**: Cada objetos se "dibuja" con distintos formatos, estos se definene a nivel general. Hay dos tipos, los básicos o "primitivos", por ejemplo: `"right": { "align" : "right" }` y los compuestos que se definen como la suma de atributos primitivos, por ejemplo: `"encabezado": [ "default_font", "bold", "color", { "bottom" : 1, "bottom_color" : "#0000FF", "text_wrap": "True", "valign": "top" }]`, en este ejemplo "default_font", "bold", "color" son formatos definidos previamente y  { "bottom" : 1, "bottom_color" : "#0000FF", "text_wrap": "True", "valign": "top" } es un primitivo definido en el momento.

* **conditional**: Formatos condicionales

Nota: Para referencia de las definiciones, ver la documentación del módulo [XlsxWriter](http://xlsxwriter.readthedocs.io/)

Construcción de la cadena Dsn según datasource
==============================================

	* SQL Server: "DRIVER={SQL Server};SERVER=<<server>>;DATABASE=<<database>>;UID=<<usuario>>;PWD=<<password>>" 


Primeros pasos
==============

Para entender el funcionamiento de esta herramienta, vamos a imaginar el siguiente escenario: Tenemos
un conjunto de servidores SQL Server y deseamos de forma automatizada generar un informe a una determinada
hora de los procesos corriendo en los mismos. Para esto contamos con un clásico stored procedure llamado
`sp_who2`, usando `autoxls` resulta muy fácil hacer esto. El primer paso es generar la definición del
proceso de exportación de datos, esto lo haremos escribiendo un archivo JSON similar a este:

```javascript
{
	"datasources": {
		"data" : {
			"data_connect_str" : "DRIVER={SQL Server};SERVER=<<server>>;DATABASE=master;UID=<<user>>;PWD=<<passw>>",
			"data_query" : "EXEC sp_who2"
		}
	},
	"files": [
		{
			"filename": "sp_who2 on <<server>>_<<Now>>.xlsx",
			"sheets": [
				{
					"name": "sp_who2 on <<server>>",
					"default_row_height" : 11.5 ,
					"objects": {
						"text": [
							{ "text" : "Resultado del sp_who ejecutado el <<Now>> en <<server>>", "format" : "encabezado_titulo", "at" : "B2" }
						],
						"text_rows": [
							{ "text" : [null,null,null,null,null], "format" : "encabezado_titulo", "at" : "C2" }
						],
						"datagrid": [
							{
								"source" : {"datasource": "data","recordset_index" : 1},
								"at" : "B3",
								"header_format": "encabezado",
								"header_height": 25,
								"freeze_header" : true,
								"datacols" : [
												[ 1, "SPID"						, 8		, "int"				, null ],
												[ 2, "Status"					, 20	, "default"			, null ],
												[ 3, "Login"					, 16	, "default"			, null ],
												[ 4, "HostName"					, 12	, "default"			, null ],
												[11, "ProgramName"				, 60	, "default"			, null ],
												[ 8, "CpuTime"					, 12	, "number"			, null ]

								],
								"autofilter_column_range" : [1,6],
								"subtotals" : [
									{"at" : "END", "format" : "subtotal_int", "total_function" : "2" , "cols_num" : [1] },
									{"at" : "END", "format" : "subtotal", "total_function" : "9" , "cols_num" : [6] }
								]
							}
						]
					}
				}
			]	
		}
	],
	"formats": {
			"default_font"		: { "font_name" : "Verdana", "font_size" : 8, "num_format" : "", "valign" : "top" },
			"right" 			: { "align" : "right" },
			"left" 				: { "align" : "left" },
			"bold" 				: { "bold" : "True" },
			"color"				: { "bg_color": "#C6EFCE" },
			"int_fmt"			: { "num_format" : "#,##0" },
			"number2_fmt"		: { "num_format" : "#,##0.00" },
			"default" 			: [ "default_font", "left" ] ,
			"encabezado_titulo"	: [ "default_font", "bold", "color"],
			"encabezado"		: [ "default_font", "bold", "color", { "bottom" : 1, "bottom_color" : "#0000FF", "text_wrap": "True", "valign": "top" }],
			"subtotal_int" 		: [ "default_font", "right", "bold", "int_fmt" ],
			"subtotal" 			: [ "default_font", "right", "bold", "number2_fmt" ],
			"number" 			: [ "default_font", "right", "number2_fmt" ],
			"int" 				: [ "default_font", "right", "int_fmt" ]
	}
}
```


Ejecución
=========

```
#!bash

uso: autoxls [-h] [-v] [-o "path"] [-n <level>] [-l file] [-f "archivo"]
             [-k '{key:value,key:value}'] [-s]
             ["archivo"]

Generación automatizada de archivos Excel (v0.9)
2014, 2015, 2016 Patricio Moracho <pmoracho@gmal.com>

argumentos posicionales:
  "archivo"                               Archivo de entrada (JSON)

argumentos opcionales:
  -h, --help                              mostrar esta ayuda y salir
  -v, --version                           show program's version number and exit
  -o "path", --outputpath "path"          Carpeta de salida dónde se almacenaran las planillas
  -n <level>, --loglevel <level>          Nivel de log (default: ninguno)
  -l file, --logfile file                 Archivo de log
  -f "archivo", --keywordfile "archivo"   Archivo de keywords del procesos
  -k '{key:value,key:value}', --keywords '{key:value,key:value}'
                                          Keywords del procesos
  -s, --start-excel                       Abrir automáticamente las planillas generadas


```

Definiciones de keywords
========================


Niveles de log
==============

Utilizar el parámetro `-n` o `--loglevel` para indicar el nivel de información que mostrará la herramienta. Por defecto el nivel es NONE, que no mustra ninguna información.

Nível		| Detalle
----------- | -------------
NONE		| No motrar ninguna información
DEBUG		| Información detallada, tipicamente análisis y debug
INFO		| Confirmación visual de lo esperado
WARNING		| Información de los eventos no esperados, pero aún la herramienta puede continuar
ERROR		| Errores, alguna funcionalidad no se puede completar
CRITICAL 	| Errores serios, el programa no puede continuar


Notas para Desarrollo
=====================

Para desarrollo de la herramienta es necesario, además de contar con el entorno de desarrollo python mencionado [aquí](../README.md), tener en cuenta la siguiente información:

* Crear el entorno de desarrollo
	* Crear el entorno virtual, de esta manera aislamos las librerías que necesitaremos sin "ensuciar" el entorno Python base, por ejemplo: `virtualenv ../venvs/autoxls
	* Activar el entorno, antes que nada hay que activar el entorno, para que los paths a Python apunten a las nuevas carpetas (Usando bash):  `source  ../venvs/autoxls/Scripts/activate`, en Windows: `../venvs/autoxls/Scripts/activate.bat`
	* Instalar librerías adicionales. 
		* [XlsxWriter](https://github.com/jmcnamara/XlsxWriter): Estupenda libreria para generar archivos Excel.
		* [pypyodbc](https://github.com/jiangwen365/pypyodbc) para la conectividada con las bases de datos: `pip install pypyodbc`
		* [pyinstaller](https://github.com/pyinstaller/pyinstaller/) solo si el objetivo final es construir un ejecutable binario, esta herramienta es bastante sencilla y rápida si bien es mucho más poderosa [Cx_freeze](https://bitbucket.org/anthony_tuininga/cx_freeze), para instalar: `pip install pyinstaller`


* Probar el autoxls
	* Activar el entorno:  `source  ../venvs/autoxls/Scripts/activate` o `../venvs/autoxls/Scripts/activate.bat`
	* Ejecuta el script principal: `python autoxls.py -h`


* Preparar EXE para distribución
	* `pyinstaller autoxls.py -y --onefile --clean`
	* El archivo final debería estar en ./dist/autoxls.exe

