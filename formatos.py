# -*- coding: utf-8 -*-
"""
# Copyright (c) 20014 Patricio Moracho <pmoracho@gmail.com>
#
# formatos.py
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
"""


class Formatos():
	"""
	Clase para mantener una lista de formatos XlsxWriter de la planilla Excel
	"""
	def __init__(self, formats=None):
		""" __init__. Inicialización de la clase contenedora de los formatos (XlsxWriter) para la planilla excel
		
		Se procesa el diccionario de formatos recibido (formats) que puede contener:
		
			* Valores "primitivos" para XlsWriter, por ej: 
				"default_font"	: { "font_name" : "Verdana", "font_size" : 8 },
				"center"		: { "align" : "center" }

			* Valores "combinados" por ej:
				"default_center": [ "default_font", "center" ]
				"encabezado"	: [ "default_font", { "bottom" : 1, "bottom_color" : "#0000FF", "text_wrap": "True", "valign": "top" }]
		
		La lista interna _formats se completará armando los formatos definitivos: "primitivos +
		"combinados" traducidos en "primitivos", por ej: 
		
				"default_font"	: { "font_name" : "Verdana", "font_size" : 8 }
				"center"		: { "align" : "center" }
				"default_center": { "font_name" : "Verdana", "font_size" : 8, "align" : "center"  }
				"encabezado"	: [ "font_name" : "Verdana", "font_size" : 8, "bottom" : 1, "bottom_color" : "#0000FF", "text_wrap": "True", "valign": "top" }

		Args:
			formats: <dict> Dicionario de formatos

		"""
		self._formats 			= {}
		self._formats_obj		= {}
		self._active_workbook 	= None

		if formats:
			# Primer cargo los formatos "primitivos" o básicos son los que ya son diccionarios
			for k,v in [(k, v) for k, v in formats.items() if isinstance(v,dict)]:
				self._formats[k] = v
			
			# Cargo los valores tipo lista que agrupan varios valores "primitivos" el orden determina el formato final
			for k,v in [(k, v) for k, v in formats.items() if isinstance(v,list)]:
				d = {}
				for atributo in v:
					if isinstance(atributo,str):
						d.update(self._formats.get(atributo,None))

					if isinstance(atributo,dict):
						# En la lista puedo tener un diciionario con un formato primitivo
						d.update(atributo)

				self._formats[k] = d

	def __str__(self):
		return ',\n'.join("{!s}={!r}".format(key, val) for (key, val) in self._formats.items())

	def set_active_workbook(self, workbook):
		"""
		Establece la planilla activa dónde se asociarán eventualmente los formatos

		Args:

			workbook: <xlsxwriter.Workbook> representa planilla Xlsx dónde se definirán los formatos
		"""
		self._active_workbook = workbook

	def clear(self):
		self._formats_obj.clear()

	def new(self, name, format_spec):
		"""
		Crea un nuevo Formato de nombre <name> y con la especificación <format_spec>

		Args:

			name:	<str> Nombre del formato

		Returns:

			<xlsxwriter.Format> Objeto que representa el formato
		
		"""

		if self._formats_obj.get(name) is None:
			self._formats[name] = format_spec

		return self.get(name)

	def get_spec(self, name):
		return self._formats.get(name)

	def get(self, name):
		"""
		Devuelve el objeto <xlsxwriter.Format> según el nombre del mismo
		
		Cuando se invoca esta rutina para un formato, la primera vez se
		crea un nuevo objeto <xlsxwriter.Format> que se incorpora a la lista interna 
		de la clase _formats_obj. Cada vez que se utilice dicho formato se devolverá 
		el objeto, de esta forma los formatos se reutilizan y solo se crean si son
		referenciados en en la planilla.
		
		Args:

			name:	<str> Nombre del formato

		Returns:

			<xlsxwriter.Format> Objeto que representa el formato

		"""
		fmt_def = self._formats.get(name)

		if fmt_def is not None:

			if self._formats_obj.get(name) is None:
				try:
					self._formats_obj[name] = self._active_workbook.add_format(fmt_def)
					return self._formats_obj[name]

				except Exception:
					print("Imposible establecer el formato: %s" % name)
					return None
			else:
				"""
				El formato ya fue ingresado se devuelve el objeto
				"""
				return self._formats_obj[name]

		return None
