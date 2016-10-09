# -*- coding: utf-8 -*-

"""
# Copyright (c) 2014 Patricio Moracho <pmoracho@gmail.com>
#
# datasource.py
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
"""

try:
    import sys
    """
    Librerias NO estandars
    """
    import pypyodbc

except ImportError as err:
    modulename = err.args[0].partition("'")[-1].rpartition("'")[0]
    print("No fue posible importar el modulo: %s" % modulename)
    sys.exit(-1)


class datasource():
    """datasource: Clase para el manejo de datos."""

    def __init__(self, connectstr, query):
        """__init__."""

        self._connectstr 	= connectstr
        self._query 		= query
        self._data			= []
        self._newdata		= []
        self._header		= []
        self._conn 			= None
        self._cur 			= None
        self.cols 			= 0
        self.rows 			= 0

    def __del__(self):
        """__del__."""
        self.close_db()

    def open_db(self):
        """open_db: open database."""
        self._conn = pypyodbc.connect(self._connectstr)
        self._cur = self._conn.cursor()

    def close_db(self):
        """close_db: close database."""
        if self._cur is not None:
            self._cur.close()
        if self._conn is not None:
            self._conn.close()

    def header(self):
        """header."""
        return self._header

    def newdata(self, rsnum=0):
        """data: retorna recordset."""

        if self._query and self._connectstr:

            if not self._newdata:

                self.open_db()
                self._cur.execute(self._query)

                while True:
                    rs = 0

                    rows = self._cur.fetchall()

                    campos = {}
                    cols = 0
                    for coldesc in self._cur.description:
                        if coldesc[0] in campos:
                            campo = "%s_%d" % (coldesc[0], cols)
                        else:
                            campo = coldesc[0]

                        campos[campo] = campo
                        self._header.append(campo)
                        cols = cols + 1

                    d = {"rows": rows, "colnames": campos}
                    self._newdata.append(d)

                    rs = rs + 1
                    if self._cur.nextset() == False:
                        break

            return self._newdata[rsnum]

        else:
            return None

    def data(self):
        """data: retorna recordset."""

        if self._query and self._connectstr:
            if not self._data:
                self.open_db()
                self._cur.execute(self._query)

                row = 0
                for fila in self._cur.fetchall():
                    self._data.append(fila)
                    row = row + 1

                cols 			= 0
                campos			= {}

                for d in self._cur.description:
                    if d[0] in campos:
                        campo = "%s_%d" % (d[0], cols)
                    else:
                        campo = d[0]

                    campos[campo] = campo
                    self._header.append(campo)
                    cols = cols + 1

                self.cols = cols
                self.rows = row

                self._cur.next()

            return self._data

        else:
            return None

if __name__ == "__main__":

    connectstr = "DRIVER={SQL Server};SERVER=momdb2test;DATABASE=contable_DB;UID=mecanus;PWD=mecanus"
    query = "select top 3 id, name from syscolumns; select top 3 id, name, crdate from sysobjects;"
    d = datasource(connectstr, query)

    print(d.newdata(0)["colnames"])
    print(d.newdata(1)["colnames"])

    sys.exit(0)
