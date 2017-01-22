# -*- coding: UTF-8 -*-
from __future__ import unicode_literals, print_function, division
from copy import copy
import common
import process_export


def copy_style_to_new(old_cell, new_cell):
    if old_cell.has_style:
        new_cell.font = copy(old_cell.font)
        new_cell.border = copy(old_cell.border)
        new_cell.fill = copy(old_cell.fill)
        new_cell.number_format = copy(old_cell.number_format)
        new_cell.protection = copy(old_cell.protection)
        new_cell.alignment = copy(old_cell.alignment)


def write_to_hybrid_worksheet(worksheet, max_row, lines):
    for i, line in enumerate(lines, start=1):
        worksheet['E{}'.format(max_row + i)] = line['da']
        worksheet['F{}'.format(max_row + i)] = line['di']
        worksheet['G{}'.format(max_row + i)] = '=G{}+F{}'.format(max_row, max_row + i)
        print('add new line: {},{},{} to {}'.format(line['da'], line['di'], '=G{}+F{}'.format(max_row, max_row + i), worksheet.title))
        copy_style_to_new(worksheet['E{}'.format(max_row)], worksheet['E{}'.format(max_row + i)])
        copy_style_to_new(worksheet['F{}'.format(max_row)], worksheet['F{}'.format(max_row + i)])
        copy_style_to_new(worksheet['G{}'.format(max_row)], worksheet['G{}'.format(max_row + i)])


def write_to_worksheet(worksheet, max_row, lines):
    for i, line in enumerate(lines, start=1):
        worksheet['A{}'.format(max_row + i)] = line['date'].strftime('%Y/%m/%d')
        worksheet['B{}'.format(max_row + i)] = line['da']
        worksheet['C{}'.format(max_row + i)] = line['di']
        worksheet['D{}'.format(max_row + i)] = '=D{}+C{}'.format(max_row, max_row + 1)
        print('add new line: {},{},{},{} to {}'.format(line['date'].strftime('%Y/%m/%d'), line['da'], line['di'], '=D{}+C{}'.format(max_row, max_row + 1), worksheet.title))
        copy_style_to_new(worksheet['A{}'.format(max_row)], worksheet['A{}'.format(max_row + i)])
        copy_style_to_new(worksheet['B{}'.format(max_row)], worksheet['B{}'.format(max_row + i)])
        copy_style_to_new(worksheet['C{}'.format(max_row)], worksheet['C{}'.format(max_row + i)])
        copy_style_to_new(worksheet['D{}'.format(max_row)], worksheet['D{}'.format(max_row + i)])


def write_to_dest(dest_workbook, data):
    for worksheet in dest_workbook.worksheets:
        parts = worksheet.title.split(' ')
        if len(parts) != len(['code_name', 'real_name', 'module', 'os']):
            continue
        code_name = parts[0]
        real_name = parts[1]
        module = parts[2]
        os = parts[3]
        if module == '混合':
            data_key_1 = process_export.get_data_key(code_name, '用户画像', os)
            data_key_2 = process_export.get_data_key(code_name, '运营统计', os)
            max_row = worksheet.max_row
            if data_key_1 in data:
                write_to_worksheet(worksheet, max_row, data[data_key_1])
            if data_key_2 in data:
                write_to_hybrid_worksheet(worksheet, max_row, data[data_key_2])
        else:
            data_key = process_export.get_data_key(code_name, module, os)
            if data_key not in data:
                continue
            write_to_worksheet(worksheet, worksheet.max_row, data[data_key])


def main():
    export_filename, dest_filename = common.get_export_and_dest_filename()
    export_workbook = common.get_workbook(export_filename, read_only=True)
    export_data = process_export.group_by_line_key(process_export.read_lines_from_worksheet(export_workbook.active))
    dest_workbook = common.get_workbook(dest_filename)
    write_to_dest(dest_workbook, export_data)
    dest_workbook.save(dest_filename)


if __name__ == '__main__':
    main()
