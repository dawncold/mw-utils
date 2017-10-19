# -*- coding: UTF-8 -*-
from __future__ import unicode_literals, print_function, division

import csv
import sys
from mmap import mmap, ACCESS_READ
from decimal import Decimal

import xlrd

WRITE_TO_SHEET_NAME = 'OBF'
DEFAULT_DATA_DAYS = 7
DEFAULT_UV_COLUMN_INDEX = 1
DEFAULT_UI_COLUMN_INDEX = 3  # user increment


def clear_sheet(sheet, data_sheet):
    for i in xrange(data_sheet.nrows):
        for j in xrange(data_sheet.ncols):
            sheet.write(i, j, label='')


def calculate_avg_and_sum(data_sheet, days):
    nrows = data_sheet.nrows
    if not data_sheet.cell(nrows - 1, 0).value:
        while not data_sheet.cell(nrows - 1, 0).value:
            nrows -= 1
            if nrows <= 0:
                raise Exception('invalid data')
        sum = int(data_sheet.cell(nrows - 1, DEFAULT_UI_COLUMN_INDEX).value)
    else:
        sum = int(data_sheet.cell(data_sheet.nrows - 1, DEFAULT_UI_COLUMN_INDEX).value)
    if nrows <= days:
        days = nrows - 1  # exclude first row as it is header
    avg_data = []
    while days > 0:
        nrows -= 1
        days -= 1
        # print(data_sheet.name, data_sheet.cell(nrows, DEFAULT_UV_COLUMN_INDEX).value)
        value = str(data_sheet.cell(nrows, DEFAULT_UV_COLUMN_INDEX).value).strip()
        if not value:
            continue
        avg_data.append(int(Decimal(value)))

    if not avg_data:
        avg = 0
    else:
        avg = int(reduce(lambda x, y: x + y, avg_data) / len(avg_data))

    return avg, sum


def main():
    file_path = sys.argv[1].decode('UTF-8')

    with open(file_path, 'rb') as f:
        book = xlrd.open_workbook(file_contents=mmap(f.fileno(), 0, access=ACCESS_READ))

    data = []

    for i in xrange(book.nsheets):
        s = book.sheet_by_index(i)
        if s.visibility != 0:
            continue
        if s.name == WRITE_TO_SHEET_NAME:
            continue
        elif s.name == '汇总':
            continue
        else:
            days = int(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_DATA_DAYS
            avg, sum = calculate_avg_and_sum(s, days)
            data.append([s.name, unicode(avg), unicode(sum)])
    header = ['TAB NAME', 'AVG', 'TOTAL']
    with open('{}.csv'.format(file_path), mode='wb+') as out_file:
        csv_writer = csv.writer(out_file)
        csv_writer.writerow(header)
        for record in data:
            csv_writer.writerow([field.encode('utf8') if field else b'' for field in record])


if __name__ == '__main__':
    main()
