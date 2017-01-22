# -*- coding: UTF-8 -*-
from __future__ import unicode_literals, print_function, division
import zipfile
import openpyxl


def get_workbook(filename, read_only=False):
    try:
        workbook = openpyxl.load_workbook(filename, read_only=read_only)
    except zipfile.BadZipfile:
        print('File is not an valid Microsoft Office xlsx file, please convert it to xlsx file!')
        exit(-1)
    else:
        return workbook
