# -*- coding: utf-8 -*-
from __future__ import unicode_literals, print_function, division
import sys
from common_openpyxl import *


def get_filename():
    arguments = sys.argv
    if len(arguments) < 2:
        print('please provide filename!')
        exit(-1)
    return sys.argv[1]


def get_export_and_dest_filename():
    arguments = sys.argv
    if len(arguments) < 3:
        print('please provide export filename and dest filename!')
        exit(-1)
    return sys.argv[1], sys.argv[2]
