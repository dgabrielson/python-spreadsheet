#!/usr/bin/env python
"""
Reading and writing ODS OpenDocument Format spreadsheets
"""
from __future__ import print_function, unicode_literals

import re
from io import StringIO

from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.style import (ParagraphProperties, Style, TableColumnProperties,
                       TextProperties)
from odf.table import Table, TableCell, TableColumn, TableRow
from odf.text import P

import spreadsheet  # for spreadsheet.Formula

PWENC = "utf-8"


def reader(filename, fileobj, **kwargs):
    """
    ``fileobj`` should be in binary, and at the beginning of the stream.
    
    """
    # load_workbook backs onto zipfile.ZipFile, which supports file objects or filenames.
    book = load(fileobj or filename) 
    sheet = book.spreadsheet
    
    results = []
    for tr in sheet.getElementsByType(TableRow):
        row = []
        for tc in tr.getElementsByType(TableCell):
            value = None
            for item in tc.getElementsByType(P):
                value = item.firstChild.data
                break   # there can be only one p per tc!
            row.append( value )
        while row[-1] is None:
            row.pop()
        results.append( row )   # this processes formulas.
    return results



def writer(data, **kwargs):
    """
    Liberally adapted from odfpy's csv2ods script.
    """
    def handle_formula(f):
        return TableCell(valuetype="float", formula="{}".format(f))
        

    textdoc = OpenDocumentSpreadsheet(**kwargs)
    # Create a style for the table content. One we can modify
    # later in the word processor.
    tablecontents = Style(name="Table Contents", family="paragraph")
    tablecontents.addElement(ParagraphProperties(numberlines="false", linenumber="0"))
    tablecontents.addElement(TextProperties(fontweight="bold"))
    textdoc.styles.addElement(tablecontents)
    
    # Start the table
    table = Table( name=u'Sheet 1' )
    fltExp = re.compile('^\s*[-+]?\d+(\.\d+)?\s*$')
    
    for row in data:
        tr = TableRow()
        table.addElement(tr)
        for val in row:
            if isinstance(val, spreadsheet.Formula):
                tc = handle_formula(val)  
            else:
                valuetype = 'string'
                if not isinstance(val, unicode):
                    if isinstance(val, str):
                        text_value = "{}".format(val, PWENC, 'replace')
                    else:
                        text_value = "{}".format(val)
                else:
                    text_value = val
    
                if isinstance(val, (float,int,long)) or ( isinstance(val, str) and fltExp.match(text_value) ):
                    valuetype = 'float'
                    
                if valuetype == 'float':  
                    tc = TableCell(valuetype="float", value=text_value.strip())
                else:
                    tc = TableCell(valuetype=valuetype)
                    
                if val is not None:
                    p = P(stylename=tablecontents,text=text_value)
                    tc.addElement(p)

            tr.addElement(tc)



    textdoc.spreadsheet.addElement(table)

    result = StringIO()
    textdoc.write( result )
    return result.getvalue()
    
    
##
