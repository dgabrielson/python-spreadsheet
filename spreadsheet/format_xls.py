#!/usr/bin/env python
"""
Reading and writing XLS Excel spreadsheets
"""
from __future__ import print_function, unicode_literals

from io import StringIO

import xlrd
import xlwt

import spreadsheet  # for spreadsheet.Formula


def reader(filename, fileobj, **kwargs):
    """
    ``fileobj`` should be in binary, and at the beginning of the stream.
    
    """
    if fileobj is not None:
        file_contents = fileobj.read()
        book = xlrd.open_workbook(file_contents=file_contents, **kwargs)
    else:
        book = xlrd.open_workbook(filename=filename, **kwargs)
    
    first_sheet_name = book.sheet_names()[0]
    sheet = book.sheet_by_name(first_sheet_name)
    results = []
    for i in range(sheet.nrows):
        results.append( sheet.row_values(i) )   # this processes formulas.
    return results




def writer(data, **kwargs):
    """
    """
    def handle_formula(f):
        return xlwt.Formula(f.formula_string)
        

    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet(u'Sheet 1')
    for i in range(len(data)):
        row = data[i]
        for j in range(len(row)):
            cell = row[j] if not isinstance(row[j], spreadsheet.Formula) else handle_formula(row[j])
            sheet.write(i,j, cell)
    
    result = StringIO()
    book.save(result)
    return result.getvalue()
    
    
##
