# -*- coding: utf-8 -*-
# Copyright (c) 20014 Patricio Moracho <pmoracho@gmail.com>
#
# engine.py
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of version 3 of the GNU General Public License
# as published by the Free Software Foundation. A copy of this license should
# be included in the file GPL-3.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Library General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.

##################################################################################################################################################
# Imports
##################################################################################################################################################
try:
	import sys
	import json
	import re
	import datetime
	import string
	import os
	import codecs
	import hashlib
	"""
	Clases propias
	"""
	from datasource import datasource
	from formatos import Formatos
	"""
	Librerias NO estandars
	"""
	import xlsxwriter
	from xlsxwriter.utility import xl_cell_to_rowcol
	from xlsxwriter.utility import xl_range

except ImportError as err:
	modulename = err.args[0].partition("'")[-1].rpartition("'")[0]
	print("No fue posible importar el modulo: %s" % modulename)
	sys.exit(-1)


def sha256(string):
	hash_object = hashlib.sha256(string.encode())
	return hash_object.hexdigest()

##################################################################################################################################################
# Clases
##################################################################################################################################################


class Engine():
	"""Engine."""

	def __init__(self, jsonfile, keywords=None, logging=None):
		"""__init__."""
		self.inputfile 			= jsonfile
		self.active_workbook 	= None
		self.active_worksheet	= None
		self.logging			= logging
		self.keywords 			= {}
		self.conditionals		= {}
		self.datasources		= {}
		self.regex				= ""
		self.now 				= datetime.datetime.now()
		
		self.fr					= 9999999
		self.fc					= 9999999
		self.er					= 0
		self.ec					= 0

		keywords["Now"] 		= self.now.strftime("%Y-%m-%d %H:%M:%S")
		self.keywords 			= dict((("<<%s>>" % key), value) for key, value in keywords.items())

		# Create a regular expression  from the dictionary keys
		self.regex 			= re.compile("|".join(self.keywords.keys()))

		try:
			with open(self.inputfile, "r", encoding='utf8') as json_file:
				self.json_data = json.load(json_file)

				# Verificar los keywords del template que no estuvieran definidos
				keys_in_template = set(re.findall(r"\<<(\w+)\>>", json.dumps(self.json_data)))
				keys_faltantes = [x for x in keys_in_template if ("<<%s>>" % x) not in self.keywords.keys()]

				if keys_faltantes != []:
					raise ValueError("Faltan definir las siguientes keywords: %s" % keys_faltantes)

				# Carga de los Formatos
				self.formatos = Formatos(self.json_data.get("formats", {}))

				# Carga de los Formatos condicionales
				self.conditionals = self.json_data.get("conditional")

				# Carga de los Data sources
				dss = self.json_data.get("datasources", [])
				for each in dss:
					properties = dss[each]
					query = ''
					if properties.get("data_query_file") is not None:
						sqlfile = properties.get("data_query_file")
						path = os.path.dirname(sqlfile)
						if path == "" or path == ".":
							path = os.path.dirname(self.inputfile)
							sqlfile = os.path.join(path, os.path.basename(sqlfile))
						try:
							query = self._read_utf8ascii_file_as_uni(sqlfile)
						except IOError as inst:
							self.error(u"Ocurrio el error %s al intentar abrir el archivo SQL '%s'" % (inst.args, sqlfile))
							sys.exit(-1)
					else:
						if type(properties["data_query"]) is list:
							query = ''.join(properties["data_query"])
						else:
							query = properties["data_query"]

					query = self.get_string_from_template(query)
					ds = datasource(self.get_string_from_template(properties["data_connect_str"]), query)
					self.datasources[each] = ds

		except IOError as inst:
			self.error(u"Ocurrio el error %s al intentar abrir el archivo de entrada '%s'" % (inst.args, self.inputfile))
			sys.exit(-1)
		except ValueError as inst:
			self.error(u"Ocurrio el error %s al intentar interpretar el archivo de entrada '%s'" % (inst.args, self.inputfile))
			sys.exit(-1)

	def _read_utf8ascii_file_as_uni(self, fname):
		"""Intenta leer un archivo como utf8 sino lo considera un ascii estándar."""

		try:
			self._inputfile_encoding = 'utf8'
			with codecs.open(fname, 'r', encoding='utf8') as f:
				return f.read()

		except UnicodeError:

			self._inputfile_encoding = 'iso-8859-1'
			with codecs.open(fname, 'r', encoding='iso-8859-1') as f:
				return f.read()

	def info(self, msg):
		if self.logging:
			self.logging.info(msg)

	def error(self, msg):
		if self.logging:
			self.logging.error(msg)

	def generate(self, outputpath, startexcel):
		"""generate. Genera los archivos Xlsx.

		Args:
			outputpath:	(string) path de output de los archivos Excel a generar
			startexcel: (bool) True si se desea inciar el Excel y abrir cada uno de los aarchivos generados

		"""
		for file in self.json_data["files"]:
			if file.get("enabled", True):
				filename = self.generate_file(file, outputpath)
				if startexcel:
					self.info("Intentando abrir: {0}".format(filename))
					try:
						os.startfile(filename)
					except FileNotFoundError as e:
						self.error(u"Ocurrio el error %s al intentar abrir el archivo '%s'" % (str(e), filename))

	def get_string_from_template(self, text):
		"""get_string_from_template: Reemplazar los "keywords" por valores reales."""
		if text:
			return self.regex.sub(lambda m: self.keywords[m.group(0)], text)
		return None

	def generate_file(self, def_file, outputpath):
		"""generate_file: Genera un archivo excel.

		Args:
			def_file: 	(dict) Definición del archivo a generar
			outputpath:	(string) Carpeta donde se salvará el archivo Excel

		Returns:

			(string) path y nombre real del archivo generado

		"""

		realfilename = os.path.join(outputpath, self._normalize_filename(self.get_string_from_template(def_file["filename"])))

		self.info("Generando {0}...".format(realfilename))

		self.active_workbook = xlsxwriter.Workbook(realfilename, {'strings_to_numbers': True})
		self.formatos.set_active_workbook(self.active_workbook)

		# Generación de las solapas "habilitadas"
		for sheet in [s for s in def_file["sheets"] if s.get("enabled") is not False]:
			self.generate_sheet(sheet)

		try:
			self.active_workbook.close()
			self.formatos.clear()

		except IOError:
			self.error("Imposible salvar el archivo %s. Estara abierto o no exsite la carpeta?" % realfilename)
		except Exception as e:
			self.error("Imposible salvar el archivo %s. Error: %s" % (realfilename, str(e)))

		return realfilename

	def generate_sheet(self, sheet):
		"""generate_sheet: Genera una solapa.

		Args:
			sheet: 	(dict) Definición de la solapa a generar

		"""

		sheet_name = self.get_string_from_template(sheet.get("name", "Hoja"))[:31]
		self.info("Solapa: {0}".format(sheet_name))
		self.active_worksheet = self.active_workbook.add_worksheet(sheet_name)
		self.active_worksheet.set_default_row(sheet.get("default_row_height", 11.5))

		objects	= sheet.get("objects", [])

		for o in objects.get("text", []):
			self.insert_text(o)

		for o in objects.get("text_formated", []):
			self.insert_text_formated(o)

		for o in objects.get("text_rows", []):
			self.insert_text_rows(o)

		try:
			for o in objects.get("table", []):
				self.insert_table(o)
		except Exception as e:
			self.error(u"Ocurrio el error %s al intentar crear los objetos [table]" % (str(e)))

		for o in objects.get("datagrid", []):
			self.insert_datagrid(o)

		for o in objects.get("formulas", []):
			self.insert_formula(o)

		for o in objects.get("text_end", []):
			self.insert_text(o)

		print_settings = sheet.get("print", None)
		if print_settings:
			self.set_print_options(print_settings)

	def cast_text(self, text, type):
		"""cast_text: Castea un string a alguno de los tipos básicos."""
		casters = {
			"datetime"	: "datetime.datetime.strptime('{0}', '%Y%m%d').date()".format(text),
			"float"		: "float({0})".format(text),
			"int"		: "int({0})".format(text)
	  	}

		value = None
		try:
			value = eval(casters[type])
		except (TypeError, ValueError):
			# Imposible castear del dato
			pass
		return value

	def insert_text_formated(self, objeto):
		"""insert_text: Inserta un texto formateado."""

		self.info("Objetos text_formated")

		values 	= [self.cast_text(self.get_string_from_template(v), t) for v, t in objeto.get("values", [])]
		at 		= objeto.get("at")
		texto 	= objeto.get("text", "").format(*values)
		format 	= objeto.get("format")
		mrange	= objeto.get("merge_range")
		if texto and at:
			if mrange:
				self.active_worksheet.merge_range(mrange, texto, self.formatos.get(format))
			else:
				self.active_worksheet.write(at, texto, self.formatos.get(format))
		else:
			if not texto and at:
				self.active_worksheet.write_blank(at, '', self.formatos.get(format))

		self._setup_boundaries_at(at)
		
	def insert_formula(self, objeto):
		"""insert_formula. Inserta una formula."""

		self.info("Objetos formulas")

		at 		= objeto.get("at", None)
		formula = objeto.get("formula", None)
		fmt 	= self.formatos.get(objeto.get("format", None))
		if at and formula and fmt:
			self.active_worksheet.write_formula(at, formula, fmt)

		self._setup_boundaries_at(at)

	def insert_text_rows(self, objeto):
		"""insert_text_rows: Inserta filas con texto."""

		self.info("Objetos text_rows")

		row, col 	= xl_cell_to_rowcol(objeto.get("at"))
		textos 		= objeto.get("text", None)
		format 		= self.formatos.get(objeto.get("format"))

		for i, t in enumerate([self.get_string_from_template(t) for t in textos], 0):
			at = xl_range(row, col + i, row, col + i)
			if t:
				self.active_worksheet.write(at, t, format)
			else:
				self.active_worksheet.write_blank(at, '', format)

			self._setup_boundaries_rc(row=row, col=col + i)


	def insert_text(self, objeto):
		"""insert_text: Inserta un texto."""

		self.info("Objetos text")

		at 		 = objeto.get("at")
		texto 	 = self.get_string_from_template(objeto.get("text", None))
		format 	 = objeto.get("format")
		mrange	 = objeto.get("merge_range")
		if texto and at:
			if mrange:
				self.active_worksheet.merge_range(mrange, texto, self.formatos.get(format))
			else:
				self.active_worksheet.write(at, texto, self.formatos.get(format))
		else:
			if not texto and at:
				self.active_worksheet.write_blank(at, '', self.formatos.get(format))

		self._setup_boundaries_at(at)

	def insert_datagrid(self, objeto):
		"""insert_datagrid: Inserta una grilla.

		Args:
			objeto: 	(dict) Definición de la grilla

		"""

		self.info("Procesando objetos datagrid...")
		source = objeto["source"]
		altcolor = objeto.get("alternate_colors")
		ds = self.datasources.get(source["datasource"])
		if ds is None:
			self.error("No se ha definido el datasource {0}".format(source["datasource"]))
		else:

			rsnum = source.get("recordset_index", 1) - 1
			data = ds.newdata(rsnum)
			"""
			HEADER
			"""
			self.info("Creando encabezados...")
			fmt_header = self.formatos.get(objeto.get("header_format"))
			fmt_header_spec = self.formatos.get_spec(objeto.get("header_format"))
			header_row, header_col = xl_cell_to_rowcol(objeto.get("at", "A1"))

			col = header_col
			row = header_row

			self._setup_boundaries_rc(row, col)

			header = objeto.get("datacols", [])
			header_height = objeto.get("header_height", 30)

			newfmt_def = {}
			newfmt_def.update(fmt_header_spec)

			"""
			Header: Titulos
			"""
			if header != []:

				for index, titulo, width, format, conditional in header:

					if format != "v|f":
						# Combino el formato de la columna con el del header para aplicar solo sobre el header
						newfmt_def.update(self.formatos.get_spec(format))
						newfmt = self.formatos.new(sha256(str(newfmt_def)), newfmt_def)
						self.active_worksheet.write(row, col, self.get_string_from_template(titulo), newfmt)
						# Configuro las columnas
						fmt = self.formatos.get(format)
						self.active_worksheet.set_column(col, col, width, fmt)
					else:
						self.active_worksheet.write(row, col, self.get_string_from_template(titulo), fmt_header)
						self.active_worksheet.set_column(col, col, width)

					self.active_worksheet.set_row(row, header_height)
					col = col + 1


			else:
				for each in ds.header():
					self.active_worksheet.write(row, col, each, fmt_header)
					col = col + 1

			self._setup_boundaries_rc(row, col-1)

			if objeto.get("freeze_header", False):
				self.active_worksheet.freeze_panes(header_row + 1, 0)  # Freeze the first row.

			"""
			Data
			"""
			self.info("Generando grilla de datos...")
			data_col 	= header_col
			data_row 	= header_row + 1
			col 		= data_col
			row  		= data_row
			for record in data["rows"]:
				for c, f in [(index, format) for index, titulo, width, format, conditional in header]:

					list_fmt = []
					if f == "v|f":
						pos = record[c-1].rfind('|')
						if pos:
							cellformat = record[c-1][pos+1:]
							cellvalue = record[c-1][:pos]
						else:
							cellformat = f
							cellvalue =  record[c-1]
					else:
						cellformat = f
						cellvalue =  record[c-1]

					if altcolor:
						color = altcolor[1] if row % 2 == 0 else altcolor[0]
						list_fmt.append(color)					

					list_fmt.append(cellformat)

					cellfmt = self.formatos.get_new_spec_from_list(list_fmt)

					fmt = self.formatos.new(sha256(str(cellfmt)), cellfmt)
					self.active_worksheet.write(row, col, cellvalue, fmt)
					
					col = col + 1


				row = row + 1
				self._setup_boundaries_rc(row-1, col-1)
				col = data_col

			total_rows	= row - data_row

			"""
			Sub Totales
			"""
			subtotales = objeto.get("subtotals")
			if subtotales is not None:
				self.info("Creando subtotales...")
				for subtotal in subtotales:
					at = subtotal.get("at")
					
					# self._setup_boundaries_at(at)

					fmt_formula = self.formatos.get(subtotal.get("format"))
					funcion = subtotal.get("total_function")
					col_num = subtotal.get("cols_num")
					for eachcol in col_num:
						rango = xl_range(data_row, data_col + eachcol - 1, row - 1, data_col + eachcol - 1)
						formula = "=SUBTOTAL(%s,%s)" % (funcion, rango)
						if at == "END":
							at = xl_range(row, data_col + eachcol - 1, row, data_col + eachcol - 1)
							self._setup_boundaries_rc(row, data_col + eachcol - 1)
							
						self.active_worksheet.write_formula(at, formula, fmt_formula, -342047.61)

			"""
			Autofiltros por columnas
			"""
			autofilter_columns = objeto.get("autofilter_column_range")
			if autofilter_columns:
				self.info("Creando autofiltros...")
				rango = xl_range(data_row - 1, data_col + autofilter_columns[0] - 1, total_rows + 1, data_col + autofilter_columns[1] - 1)
				self.active_worksheet.autofilter(rango)

			"""
			Formatos condicionales
			"""
			col = data_col
			for (index, titulo, width, format, conditional) in header:
				if conditional is not None:
					formula = self.conditionals.get(conditional)
					if formula is not None:
						for each in formula:
							cell = xl_range(data_row, col, row, col)
							fspec = each.get("format")
							filtered_dict = {k: v for (k, v) in each.items() if "format" not in k}
							if fspec is not None:
								fmt = self.formatos.get(fspec)
								filtered_dict.update({"format": fmt})
								self.active_worksheet.conditional_format(cell, filtered_dict)
				col = col + 1

		# self._setup_boundaries(row=row, col=col)


	def insert_table(self, objeto):
		"""insert_table: Inserta una tabla."""

		self.info("Objetos table")

		source = objeto["source"]
		ds = self.datasources.get(source["datasource"])

		if ds is None:
			self.error("No se ha definido el datasource {0}".format(source["datasource"]))
		else:

			rsnum = source.get("recordset_index", 1) - 1
			data = ds.newdata(rsnum)
			row, col = xl_cell_to_rowcol(objeto.get("at"))
			self._setup_boundaries_rc(row, col)

			fmt_header = self.formatos.get(objeto.get("header_format"))

			header = []
			for each in data["colnames"]:
				header.append({"header": each, "header_format": fmt_header})

			self.active_worksheet.add_table(row,
											col,
											row + len(data["rows"]),
											col + len(data["colnames"]) - 1,
											{
												'data': data["rows"],
												'style': objeto.get("style", "Table Style Medium 2"),
												'total_row': objeto.get("total_row", 1),
												'autofilter': objeto.get("autofilter", False),
												'columns': header,
											})

			self._setup_boundaries_rc(row + len(data["rows"]), col + len(data["colnames"]) - 1)
											
	def set_print_options(self, objeto):
		"""set_print_options: Configura opciones de impresión."""

		self.info("Configuración de opciones de impresión")
		
		self.active_worksheet.set_paper(objeto.get("paper", 0))
		
		l, r, t, b = objeto.get("margins", [0.7,0.7,0.75,0.75])
		self.active_worksheet.set_margins(l, r, t, b)

		t, o = objeto.get("header", ["", None])
		t = self.get_string_from_template(t)
		self.active_worksheet.set_header(t, o)

		t, o = objeto.get("footer", ["", None])
		t = self.get_string_from_template(t)
		self.active_worksheet.set_footer(t, o)

		if objeto.get("landscape", False):
			self.active_worksheet.set_landscape()

		if objeto.get("grid", False):
			self.active_worksheet.hide_gridlines()

		if objeto.get("center_horizontally", False):
			self.active_worksheet.center_horizontally()

		if objeto.get("center_vertically", False):
			self.active_worksheet.center_vertically()

		pa = objeto.get("area", "auto")
		if pa:
			if isinstance(pa, str):
				if pa == "auto":
					fr, fc, er, ec = (self.fr, self.fc, self.er, self.ec)

			else:
				if isinstance(pa, list):
					fr, fc, er, ec = pa

		self.active_worksheet.print_area(fr, fc, er, ec)

		w, h = objeto.get("fit_to_pages", [1, 1])
		self.active_worksheet.fit_to_pages(w, h)

		self.active_worksheet.set_print_scale(objeto.get("scale", 100))

		pass
													

	def _normalize_filename(self, filename):
		"""_normalize_filename: Generates an slightly worse ASCII-only slug.

		Args:
			filename: 	(str) Nombre del archivo

		Return:
			(str) Nombre válido de archivo

		"""
		valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
		return ''.join(c for c in filename if c in valid_chars)

	def _setup_boundaries_at(self, at):

		row, col 	= xl_cell_to_rowcol(at)
		self._setup_boundaries_rc(row, col)

	def _setup_boundaries_rc(self, row, col):

		self.fr		= row if row < self.fr else self.fr
		self.fc		= col if col < self.fc else self.fc
		self.er		= row if row > self.er else self.er
		self.ec		= col if col > self.ec else self.ec
		
