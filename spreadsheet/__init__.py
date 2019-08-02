#!/usr/bin/env python
"""
This is a library for reading and writing spreadsheets.

from spreadsheet import sheetReader, sheetWriter

``sheetReadersheetReader(filename=None, fileobj=None, input_format=None, subargs={}) 
        -> <list of <list of cells>>``

``sheetWriter(data, output_format, subargs={}) -> <data string>``

These functions will only read and write the first sheet, for spreadsheet formats
that support more than one.

TODO: Add formula abstraction class.

"""
from __future__ import unicode_literals, print_function


SUPPORTED_FORMATS = [
    'csv', 'txt',   # handled by format_csv
    'xls',          # handled by format_xls
    'xlsx', 'xlsm', # handled by format_xlsx
    'odf', 'ods',   # handled by format_odf
    ]


import os

from . import (format_csv, format_xls, format_xlsx, format_odf,
               )

# why reinvent the wheel...?
from openpyxl.cell import get_column_letter, coordinate_from_string, column_index_from_string




def rowcol2coord(row, col):
    """
    Transform the 0-based row, column index into a spreadsheet style coordinate 
    e.g., maps 0,0 -> A1 ; 3,12 -> N3
    """
    return '%s%s' % ( get_column_letter(col+1), row+1 )
    
    
    
    
def coord2rowcol(coord):
    """
    Transform a spreadsheet style coordinate into a 0, 0 based index.
    """
    col, row_idx = coordinate_from_string(coord)
    col_idx = column_index_from_string(col)
    return row_idx-1, col_idx-1
    



class Formula:
    """
    A generic formula class.  Use this when a data item should be interpreted as a forumla
    in your spreadsheet output via ``sheetWriter()``.
    
    This is mainly here for type checking... each submodule is responsible for correctly
    processing the formula.
    
    Be exceptionally cautious with formulas, as supported formulas vary from program to program.
    
    In particular, formulas in CSV files render as strings (with a leading '=' sign).
    (Maybe eventually I'll write a simple parser for formulas that the format_csv module will
     evaluate itself.  Maybe.)
     
    Formulas should be specified using spreadsheet coordinates, not row,column indexes.
     
    """
    def __init__(self, formula_string=''):
        """
        ``formula_string`` should not include a leading '=', as used in many spreadsheets.
        """
        self.formula_string = "{}".format(formula_string)
        
    def __str__(self):
        return "{}".format(self.formula_string)
        
    def __repr__(self):
        return u'spreadsheet.Formula(%r)' % self.formula_string
        
        


def sheetReader(filename=None, fileobj=None, input_format=None, subargs={}):
    """
    Spreadsheet read dispatcher.
    
    Note that, in the case where the data is already around, it would be necessary to 
    buffer the data a file-like object and pass that in as a ``fileobj``.
    
    Cannot have ``filename`` and ``fileobj`` both None.
    Cannot have ``filename`` and ``format`` both None.
    
    ``fileobj`` will be used before ``filename``.
    ``format`` will be used before ``filename``.
    
    ``subargs``, if given, will be passed down to the dispatched reader.
    """
    assert filename or fileobj is not None, 'You must specify at least one of filename or fileobj'
    assert filename or input_format is not None, 'You must specify at least one of filename or format'
    if input_format is None:
        # determine format from filename
        input_format = os.path.splitext(filename)[-1].lstrip('.')
    
    assert input_format in SUPPORTED_FORMATS, 'The format %r is not supported' % input_format

    if input_format in ['csv', 'txt', ]:
        input = fileobj or filename
        return format_csv.reader(input, **subargs)
        
    if input_format in ['xls', ]:
        return format_xls.reader(filename, fileobj, **subargs)
        
    if input_format in ['xlsx', 'xlsm', ]:
        return format_xlsx.reader(filename, fileobj, **subargs)
        
    if input_format in ['odf', 'ods', ]:
        return format_odf.reader(filename, fileobj, **subargs)
        
    assert False, 'Unsupported file format -- this was already checked for and should never happen.'
    
    
    
    
    
def sheetWriter(data, output_format, subargs={}):
    """
    Spreadsheet write dispatcher.
    
    This function always returns a data stream containing the formatted contents of the 
    spreadsheet.  It makes no assumptions about wanting the format on disk or over the network.
    """
    assert output_format in SUPPORTED_FORMATS, 'The format %r is not supported' % output_format

    if output_format in ['csv', 'txt']:
        return format_csv.writer(data, **subargs)

    if output_format in ['xls', ]:
        return format_xls.writer(data, **subargs)
        
    if output_format in ['xlsx', 'xlsm', ]:
        return format_xlsx.writer(data, **subargs)
        
    if output_format in ['odf', 'ods', ]:
        return format_odf.writer(data, **subargs)
        
    assert False, 'Unsupported file format -- this was already checked for and should never happen.'



#