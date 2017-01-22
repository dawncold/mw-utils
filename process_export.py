# -*- coding: UTF-8 -*-
from __future__ import unicode_literals, print_function, division
import sys
from datetime import datetime
from pprint import pprint
import common

RETRIEVE_COLUMNS = {
    0: ('date', lambda v: datetime.strptime(v, '%Y%m%d').date()),
    1: ('name', lambda v: v),
    2: ('key', lambda v: v),
    3: ('module', lambda v: v),
    4: ('os', lambda v: v.lower()),
    5: ('da', int),
    6: ('di', int)
}


def get_data_key(code_name, module, os):
    return '{}:{}@{}'.format(code_name, module, os)


def get_filename():
    arguments = sys.argv
    if len(arguments) < 2:
        print('please provide filename!')
        exit(-1)
    return sys.argv[1]


def read_lines_from_worksheet(worksheet):
    def is_first_row(row):
        return row[1].value == 'App_Name' or row[2].value == 'App_Key'

    lines = []
    for row in worksheet.rows:
        if is_first_row(row):
            continue
        line = dict(
            (name_and_normalizer[0], name_and_normalizer[1](row[column_index].value))
            for column_index, name_and_normalizer in RETRIEVE_COLUMNS.items()
        )
        lines.append(line)
    return lines


def group_by_line_key(lines):
    data = {}
    for line in lines:
        data.setdefault(get_data_key(line['name'], line['module'], line['os']), []).append(line)
    for lines in data.values():
        lines.sort(key=lambda v: v['date'])
    return data


def main():
    filename = get_filename()
    workbook = common.get_workbook(filename)
    print(workbook.get_sheet_names())
    worksheets = workbook.worksheets
    assert len(worksheets) == 1, 'support only one worksheet'
    lines = read_lines_from_worksheet(worksheets[0])
    app_data = group_by_line_key(lines)
    for line in app_data.values()[5]:
        pprint(line)

if __name__ == '__main__':
    main()
