#! /usr/bin/env python

import sys
import xlrd
import unicodecsv
from datetime import datetime

# List any datetime columns, indexed by 0.
datecolumns=[9]

book = xlrd.open_workbook(sys.argv[1])

# Assuming the first sheet is of interest
sheet = book.sheet_by_index(0)
#sheet = book.sheet_by_name('Sheet1')

csvfile = open(sys.argv[1] + '.csv', 'wb')

wr = unicodecsv.writer(csvfile, quoting=unicodecsv.QUOTE_ALL)
wr.writerow(sheet.row_values(0))

for rownum in range(1,sheet.nrows):
  for datei in datecolumns:
    if sheet.row_values(rownum)[datei] and isinstance(sheet.row_values(rownum)[datei], float):
      year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(int(sheet.row_values(rownum)[datei]), book.datemode)
      py_date = datetime(year, month, day, hour, minute)
      sheet.put_cell(rownum, datei, 1, py_date, 0)
  wr.writerow(sheet.row_values(rownum)[0:])

csvfile.close()
