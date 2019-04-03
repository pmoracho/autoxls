Autoxls
=======

* [Página del proyecto](https://pmoracho.github.io/autoxls)
* [Proyecto en github](https://github.com/pmoracho/autoxls)
* [Descarga de ejecutable para windows](https://github.com/pmoracho/autoxls/raw/master/dist/autoxls-20190403.zip)

`autoxls` es una herramienta de linea de comandos para automatizar las
generación de archivos Excel 2003 a partir de resultados obtenidos de consultas
a bases de datos. Esta basada en el uso de la excelente librería
[XlsxWriter](https://github.com/jmcnamara/XlsxWriter), consultar la completa
documentación de esta librería para más detalle

# Empecemos

Antes que nada, necesitaremos:

* [Git for Windows](https://git-scm.com/download/win) instalado y funcionando
* Una terminal de Windows, puede ser `cmd.exe`

Con **Git** instalado, desde la línea de comando y con una carpeta dónde
alojaremos este proyecto, por ejemplo `c:\proyectos`, simplemente:

``` 
c:\> c: 
c:\> cd \proyectos 
c:\> git clone <url del repositorio>
c:\> cd <carpeta del repositorio>
``` 

|                       |                                           |
| --------------------- |-------------------------------------------|
| Repositorio           | https://github.com/pmoracho/autoxls.git |
| Carpeta del proyecto  | .                                         |

## Instalación de **Python**

Para desarrollo de la herramienta es necesario, en primer término, descargar un
interprete Python. **xls2table** ha sido desarrollado con la versión 3.6, no es
mala idea usar esta versión, sin embargo debiera funcionar perfectamente bien
con cualquier versión de la rama 3x.

**Importante:** Si bien solo detallamos el procedimiento para entornos
**Windows**, el proyecto es totalmente compatible con **Linux**

* [Python 3.6.6 (32 bits)](https://www.python.org/ftp/python/3.6.6/python-3.6.6.exe)
* [Python 3.6.6 (64 bits)](https://www.python.org/ftp/python/3.6.6/python-3.6.6-amd64.exe)

Se descarga y se instala en el sistema el interprete **Python** deseado. A
partir de ahora trabajaremos en una terminal de Windows (`cmd.exe`). Para
verificar la correcta instalación, en particular que el interprete este en el `PATH`
del sistemas, simplemente corremos `python --version`, la salida deberá
coincidir con la versión instalada 

Es conveniente pero no mandatorio hacer upgrade de la herramienta pip: `python
-m pip install --upgrade pip`

## Instalación de `Virtualenv`

[Virutalenv](https://virtualenv.pypa.io/en/stable/). Es la herramienta estándar
para crear entornos "aislados" de **Python**. En nuestro ejemplo **xls2table**,
requiere de Python 3x y de varios "paquetes" adicionales de versiones
específicas. Para no tener conflictos de desarrollo lo que haremos mediante
esta herramienta es crear un "entorno virtual" en una carpeta del proyecto (que
llamaremos `venv`), dónde una vez "activado" dicho entorno podremos instalarle
los paquetes que requiere el proyecto. Este "entorno virtual" contendrá una
copia completa de **Python** y los paquetes mencionados, al activarlo se
modifica el `PATH` al `python.exe` que ahora apuntará a nuestra carpeta del
entorno y nuestras propias librerías, evitando cualquier tipo de conflicto con un
entorno **Python** ya existente. La instalación de `virtualenv` se hará
mediante:

```
c:\..\> pip install virtualenv
```

## Creación y activación del entorno virtual

La creación de nuestro entorno virtual se realizará mediante el comando:

```
C:\..\>  virtualenv venv --clear --prompt=[autoxls] --no-wheel
```

Para "activar" el entorno simplemente hay que correr el script de activación
que se encontrará en la carpeta `.\venv\Scripts` (en linux sería `./venv/bin`)

```
C:\..\>  .\venv\Scripts\activate.bat
[autoxls] C:\..\> 
```

Como se puede notar se ha cambiado el `prompt` con la indicación del entorno
virtual activo, esto es importante para no confundir entornos si trabajamos con
múltiples proyecto **Python** al mismo tiempo.

## Instalación de requerimientos

Mencionábamos que este proyecto requiere varios paquetes adicionales, la lista
completa está definida en el archivo `requirements.txt` para instalarlos en
nuestro entorno virtual, simplemente:

```
[autoxls] C:\..\> pip install -r requirements.txt
```

## Desarrollo

Si todos los pasos anteriores fueron exitosos, podríamos verificar si la
aplicación funciona correctamente mediante:

```
[autoxls] C:\..\> python autoxls.py
uso: autoxls [-h] [-v] [-o "path"] [-n <level>] [-l file] [-f "archivo"]
             [-k {'key':'value','key':'value'}] [-s] [-d]
             ["archivo"]
autoxls: error: debe indicar el archivo de input (--inputfile)

```

La ejecución sin parámetros arrojará la ayuda de la aplicación. A partir de
aquí podríamos empezar con la etapa de desarrollo.

## Generación del paquete para deploy

Para distribuir la aplicación en entornos **Windows** nos apoyaremos en
**Pyinstaller**, un modulo, instalado junto a los requerimientos, que nos
permite crear una carpeta de distribución de la aplicación totalmente portable.
Simplemente deberemos ejecutar el archivo `windist.bat`, al finalizar el
procesos deberías contar con una carpeta en `.\dist\autoxls` la cual será una
instalación totalmente portable de la herramienta, no haría falta nada más que
copiar la misma al equipo o servidor desde dónde deseamos ejecutarla.


* Preparar EXE para distribución
	* `pyinstaller autoxls.py -y --onefile --noupx`
	* El archivo final debería estar en ./dist/autoxls.exe

# Documentación

# Características principales

* Múltiples conexiones a bases de datos
* Múltiples consultas/querys/Stored procedures
* Capacidad de procesar múltiples recordsets a partir de una única consulta
* Generáción automatizada de uno o más archivos Excel por ejecución
* Los archivos generados pueden salvarse en:
  * Una carpeta definida
  * El escritorio del usuario `{Desktop}`
  * Una carpeta temporal `{Tmp}`
  * Abrirse automáticamente
* Definición dinámica de nombres y textos a partir de "keywords" definidas en
  la invocación o por un archivo externo de keywords, puede aplicar:
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
  * Autofiltros
* Definir formatos de los objetos
  * anchos de columnas
  * colores
  * tipos de letra / tamaño
  * Alineaciones
  * formatos númericos
  * Formatos condicionales

# Antes de empezar

Para comenzar, antes de probar la ejecución de esta herramienta, es necesario
escribir un archivo de definición de la exportación a realizar. Este no es más
que un archivo en formato JSON, que describe datos, archivos, planillas y
formatos de la exportación. Nota: para validar el formato del mismo: [JSON
Editor Online](http://www.jsoneditoronline.org/). 

Este archivo definirá los siguientes elementos:

* **datasources**: uno o más conexiones a bases de datos, con las siguientes propiedades:

  * **data\_connect\_str**: que establece una conexión a un servidor de
    datos.
  * **data\_query**: que define la correspondiente consulta. Podrá ser una
    "query" común o directamente la ejecución de un stored procedure.
  * **data\_query\_file**: Archivo dónde encontramos la consulta SQL.
  
* **files**:  Que define la generación de uno o más archivos Excel. Por cada archivo se podrá definir una o más:
  
  * **sheets**: Es decir, solapas de la planilla, por cada una de estas se pueden definir varios objetos Excel:
    
    * **text**: Texto estático, normalmente titulos, se define la celda, el
      texto y el formato.
    * **text_rows**: Define una lista de textos estáticos, que se escriben
      a partir de una posición en la fila, una celda a continuación de la
      otra con un determinado formato
    * **text_formated**: Una especialización de los objetos de texto, que
      permite aplicar formatos sobre datos recibidos en las __keywords__.
    * **datagrid**: Una grilla de datos, la salida final de los datos
      recuperados.
    * **formulas**: Formulas de escel
    * **table**: Una tabla Excel. 

* **formats**: Cada objetos se "dibuja" con distintos formatos, estos se
  definene a nivel general. Hay dos tipos, los básicos o "primitivos", por
  ejemplo: `"right": { "align" : "right" }` y los compuestos que se definen
  como la suma de atributos primitivos, por ejemplo: `"encabezado": [
  "default_font", "bold", "color", { "bottom" : 1, "bottom_color" : "#0000FF",
  "text_wrap": "True", "valign": "top" }]`, en este ejemplo "default_font",
  "bold", "color" son formatos definidos previamente y  { "bottom" : 1,
  "bottom_color" : "#0000FF", "text_wrap": "True", "valign": "top" } es un
  primitivo definido en el momento.

* **conditional**: Formatos condicionales

Nota: Para referencia de las definiciones, ver la documentación del módulo [XlsxWriter](http://xlsxwriter.readthedocs.io/)

# Construcción de la cadena Dsn según datasource

  * SQL Server: "DRIVER={SQL Server};SERVER=<<server>>;DATABASE=<<database>>;UID=<<usuario>>;PWD=<<password>>" 


# Primeros pasos

Para entender el funcionamiento de esta herramienta, vamos a imaginar el
siguiente escenario: Tenemos un servidor SQL Server y deseamos de
forma automatizada generar un informe a una determinada hora de los procesos
corriendo en el mismo. Para esto contamos con un clásico stored procedure
llamado `sp_who2`, usando `autoxls` resulta muy fácil hacer esto. El primer
paso es generar la definición del proceso de exportación de datos, esto lo
haremos escribiendo un archivo JSON similar a este:

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
                "alternate_colors": ["color_impar","color_par"],
                "datacols" : [
                        [ 1, "SPID"            , 8   , "int"          , null ],
                        [ 2, "Status"          , 20  , "default"      , null ],
                        [ 3, "Login"           , 16  , "default"      , null ],
                        [ 4, "HostName"        , 12  , "default"      , null ],
                        [11, "ProgramName"     , 60  , "default"      , null ],
                        [ 8, "CpuTime"         , 12  , "number"       , "cpu" ]

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
      "default_font":      { "font_name" : "Verdana", "font_size" : 8, "num_format" : "", "valign" : "top" },
      "right":             { "align" : "right" },
      "left":              { "align" : "left" },
      "bold":              { "bold" : "True" },
      "color":             { "bg_color": "#C6EFCE" },
      "color_impar":       { "bg_color": "#A6EFCE" },
      "color_par":         { "bg_color": "#C6EFCE" },
      "int_fmt":           { "num_format" : "#,##0" },
      "number2_fmt":       { "num_format" : "#,##0.00" },
      "default":           [ "default_font", "left" ] ,
      "encabezado_titulo": [ "default_font", "bold", "color"],
      "encabezado":        [ "default_font", "bold", "color", { "bottom" : 1, "bottom_color" : "#0000FF", "text_wrap": "True", "valign": "top" }],
      "subtotal_int":      [ "default_font", "right", "bold", "int_fmt" ],
      "subtotal":          [ "default_font", "right", "bold", "number2_fmt" ],
      "number":            [ "default_font", "right", "number2_fmt" ],
      "int":               [ "default_font", "right", "int_fmt" ],
      "numero_rojo":       [ "number2_fmt", { "bg_color": "#FF0000", "font_color": "#FFFFFF"}]
  },
  "conditional": { 
    "cpu" : [
        {"type": "cell", "criteria": ">", "value": 2000, "format" : "numero_rojo" }
        ]
    }    
}

``` 

Lo llamaremos `export.json` pero puede ser cualquier nombre. Para generar
el excel a partir de la anterior definición, deberemos además establecer los
keywords del proceso:

* `<<server>>` El servidor de base de datos
* `<<user>>` Usuario
* `<<passw>>` Contraseña

Hay dos formas de hacer esto, mediante un archivo de keywords, que llamaremos
`keywords.json`, pero puede ser cualquier nombre, un texto Ascii estándar con
el siguiente formato:

```
{ 	
	"server" 			: "servidor",
	"user" 				: "miusuario",
	"passw"				: "micontraseña"
}
```

En cuyo caso ejecutaríamos así la herramienta así:

`autoxls export.json -f keywords.json`


O bien, se puede definir los "keywords" por línea de comando sin necesidad de
preconfiguraralos en un archivo así:

`autoxls export.json -k "{'server': 'servidor', 'user': 'miusuario', 'passw': 'micontraseña'}"`


# Ejecución

```bash

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
  -d, --drop-config-files                 Eliminar el archivo de input y eventualmente el de keywords

```


# Niveles de log

Utilizar el parámetro `-n` o `--loglevel` para indicar el nivel de información
que mostrará la herramienta. Por defecto el nivel es NONE, que no mustra
ninguna información.

Nível       | Detalle
----------- | -------------
NONE        | No motrar ninguna información
DEBUG       | Información detallada, tipicamente análisis y debug
INFO        | Confirmación visual de lo esperado
WARNING     | Información de los eventos no esperados, pero aún la herramienta puede continuar
ERROR       | Errores, alguna funcionalidad no se puede completar
CRITICAL    | Errores serios, el programa no puede continuar

# Notas para el desarrollador:
## Change Log:

#### Version 1.0.1 - 2019-03-29
* Se agrega nueva modalidad de formateo desde los datos : `Valor|Formato`

#### Version 1.0.1 - 2017-01-01
* Fix en el objeto "Table" y se mantiene orden original de los campos de la tabla
    

