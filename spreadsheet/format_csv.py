#!/usr/bin/env python
# -*- encoding: utf-8 -*-
"""
A module for loading and storing CSV files.
"""
from __future__ import print_function, unicode_literals

import csv
import os
from io import StringIO

try:
    unicode
except NameError:
    # Python3:
    unicode_csv_reader = csv.reader
else:
    # Python2:
    # Straight from https://docs.python.org/2/library/csv.html#csv-examples
    #   This is a python2/unicode hack.
    def unicode_csv_reader(unicode_csv_data, dialect=csv.excel, **kwargs):
        # csv.py doesn't do Unicode; encode temporarily as UTF-8:
        csv_reader = csv.reader(utf_8_encoder(unicode_csv_data),
                                dialect=dialect, **kwargs)
        for row in csv_reader:
            # decode UTF-8 back to Unicode, cell by cell:
            yield ["{}".format(cell, 'utf-8') for cell in row]

    def utf_8_encoder(unicode_csv_data):
        for line in unicode_csv_data:
            yield line.encode('utf-8')



def reader(input, delimiter=None, skip_blanks=True, **kwargs):
    """
    This function is designed to be 'the' csv loading code.

    ``input`` is either a filename, a data string, or a file-like object.
    """
    fp = None
    if callable(getattr(input, 'read', None)):
        fp = input
    else:
        fp = open(input, 'r')

    # warning: skip_blanks set False may interfer!
    if fp is not None:
        txt = fp.read()
    else:
        txt = input

    # ensure input is unicode
    if isinstance(txt, bytes):
        encoding = kwargs.pop('encoding', 'utf-8')
        txt = txt.decode(encoding, errors='replace')

    # warning: skip_blanks set False may interfer!
    txt = txt.replace('\r', '\n')


    if delimiter is None:
        # autodetect delimiter: tabs, commas, semicolons
        possible_delimiters = [',', '\t', ';', ]
        count_dict = {}
        for d in possible_delimiters:
            count_dict[d] = txt.count(d)
        delimiter = ''
        score = -1
        for d in possible_delimiters:
            if count_dict[d] > score:
                score = count_dict[d]
                delimiter = d

    results = []
    if isinstance(delimiter, bytes):
        delimiter = delimiter.encode('utf-8')
    for row in unicode_csv_reader(txt.split('\n'), delimiter=delimiter, **kwargs):
        if row or not skip_blanks:
            results.append(row)

    return results



def writer(data, **kwargs):
    """
    Improperly coded data may result in UnicodeError's.
    """
    result = StringIO()
    csv_writer = csv.writer(result, **kwargs)
    for row in data:
        csv_writer.writerow([ "{}".format(e) for e in row ])
    # Add the silly utf-8 BOM to the data stream.
    return result.getvalue().encode('utf-8-sig')

#
