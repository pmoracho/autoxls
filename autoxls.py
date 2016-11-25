# -*- coding: utf-8 -*-
"""
# Copyright (c) 2014 Patricio Moracho <pmoracho@gmail.com>
#
# autoxls.
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of version 3 of the GNU General Public License
# as published by the Free Software Foundation. A copy of this license should
# be included in the file GPL-3.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.	See the
# GNU Library General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
"""
__author__		= "Patricio Moracho <pmoracho@gmal.com>"
__appname__		= "autoxls"
__appdesc__		= "Generación automatizada de archivos Excel"
__license__		= 'GPL v3'
__copyright__	= "2014, 2015, 2016 %s" % (__author__)
__version__		= "0.9"
__date__		= "2014/11/03 13:42:03"

"""
###############################################################################
## To do´s
###############################################################################

+ Bordes de columna completa
+ Negrita los subtotales finales
+ agrupar celdas
+ output al escritorio
+ Abrir automaticamente
"""

"""
###############################################################################
# Imports
###############################################################################
"""
try:
	import sys
	import os
	import gettext
	import json
	import logging
	import tempfile

	"""
	Clases propias
	"""
	from engine import Engine

	"""
	Librerias NO estandars
	"""

	"""
	Librerias adicionales
	"""

	def my_gettext(s):
		"""my_gettext: Traducir algunas cadenas de argparse."""
		current_dict = {'usage: ': 'uso: ',
						'optional arguments': 'argumentos opcionales',
						'show this help message and exit': 'mostrar esta ayuda y salir',
						'positional arguments': 'argumentos posicionales',
						'the following arguments are required: %s': 'los siguientes argumentos son requeridos: %s'}

		if s in current_dict:
			return current_dict[s]
		return s

	gettext.gettext = my_gettext

	import argparse

except ImportError as err:
	modulename = err.args[0].split()[3]
	print("No fue posible importar el modulo: %s" % modulename)
	sys.exit(-1)

def delete_file(filename):
	try:
		os.remove(filename)
	except OSError:
		pass

def init_argparse():
	"""init_argparse: Inicializar parametros del programa."""
	usage			= "\n\n "

	cmdparser = argparse.ArgumentParser(prog			= __appname__,
										description		= "%s (v%s)\n%s\n" % (__appdesc__,__version__,__copyright__ ),
										epilog			= usage,
										formatter_class = lambda prog: argparse.RawTextHelpFormatter(prog,max_help_position=42),
										usage			= None
										)

	cmdparser.add_argument('inputfile'				, type=str, nargs='?'											, help="Archivo de entrada (JSON)", metavar="\"archivo\"")

	cmdparser.add_argument('-v', '--version'					, action='version', version=__version__)
	cmdparser.add_argument('-o', '--outputpath'		, type=str	, action="store", dest="outputpath"					, help="Carpeta de salida dónde se almacenaran las planillas", metavar="\"path\"", default=".")
	cmdparser.add_argument('-n', '--loglevel'		, type=str	, action="store", dest="loglevel"					, help="Nivel de log (default: ninguno)", metavar="<level>", default="WARNING")
	cmdparser.add_argument('-l', '--logfile'		, type=str	, action="store", dest="logfile"					, help="Archivo de log", metavar="file", default=None)
	cmdparser.add_argument('-f', '--keywordfile'	, type=str	, action="store", dest="keyworfilejson"				, help="Archivo de keywords del procesos", metavar="\"archivo\"")
	cmdparser.add_argument('-k', '--keywords'		, type=str	, action="store", dest="keyworjson"					, help="Keywords del procesos", metavar="""{'key':'value','key':'value'}""")
	cmdparser.add_argument('-s', '--start-excel'				, action="store_true", dest="startexcel"			, help="Abrir automáticamente las planillas generadas")
	cmdparser.add_argument('-d', '--drop-config-files'			, action="store_true", dest="dropcfgfiles"			, help="Eliminar el archivo de input y eventualmente el de keywords")

	return cmdparser


def file_accessible(filepath, mode):
	"""Check if a file exists and is accessible. """
	try:
		with open(filepath, mode, encoding='utf8'):
			pass
	except IOError:
		return False

	return True

"""
##################################################################################################################################################
# Main program
##################################################################################################################################################
"""
if __name__ == "__main__":

	# determine if application is a script file or frozen exe
	if getattr(sys, 'frozen', False):
		application_path = os.path.dirname(sys.executable)
	elif __file__:
		application_path = os.path.dirname(__file__)

	defaul_json_file = os.path.join(application_path, 'autoxls.json')
	defaul_keywords_file = os.path.join(application_path, 'autoxls.keywords.json')

	cmdparser = init_argparse()
	try:
		args = cmdparser.parse_args()
	except IOError as msg:
		cmdparser.error(str(msg))
		sys.exit(-1)

	# inputfile, Archivo json de definición del export
	if not args.inputfile:
		if file_accessible(defaul_json_file, 'r'):
			# Si no se pasa el archivo, y existe el autoxls.json se trata de una ejecución automática
			args.inputfile = defaul_json_file
			args.startexcel = True
			args.logfile = 'autoxls.log'
		else:
			cmdparser.error(u"debe indicar el archivo de input (--inputfile)")
			sys.exit(-1)

	log_level = getattr(logging, args.loglevel.upper(), None)
	logging.basicConfig(filename=args.logfile, level=log_level, format='%(asctime)s:%(levelname)s:%(message)s', datefmt='%Y/%m/%d %I:%M:%S', filemode='w')

	# keywords, formato python: '{key:value,key:value}'
	# Se pueden pasar por --keywords o en un archivo por --keywordfile
	# Se pueden usar luego en el Json reemplazando por el valor [keyword]
	keywords = {}
	if args.keyworjson:
		keywords.update(json.loads(args.keyworjson.replace("'", '"')))

	if not args.keyworfilejson:
		if file_accessible(defaul_keywords_file, 'r'):
			args.keyworfilejson = defaul_keywords_file

	if args.keyworfilejson:
		try:
			with open(args.keyworfilejson, "r", encoding='utf8') as json_file:
				keywords.update(json.load(json_file))

		except IOError:
			logging.error(u"Error al intentar abrir el archivo de keyords '%s'" % args.keyworfilejson)
			sys.exit(-1)

	if args.outputpath == '{desktop}':
		outputpath = os.path.join(os.path.expanduser('~'), 'Desktop')
	else:
		if args.outputpath == '{tmp}':
			outputpath = tempfile._get_default_tempdir()
		else:
			outputpath = args.outputpath

	jsonfile = args.inputfile

	logging.info("Input file  : {0}".format(jsonfile))
	logging.info("Output path : {0}".format(outputpath))
	logging.info("Keyword file: {0}".format(args.keyworfilejson))

	engine = Engine(jsonfile, keywords, logging)

	try:
		engine.generate(outputpath, args.startexcel)
		if args.dropcfgfiles:
			delete_file(jsonfile)
			delete_file(args.keyworfilejson)

	except Exception as e:
		logging.error("%s error: %s" % (__appname__, str(e)))


	logging.info("proceso exitoso!")

	sys.exit(0)
