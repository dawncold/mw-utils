# -*- coding: UTF-8 -*-
from __future__ import unicode_literals, print_function, division
import re
import common


def normalize_sheet_names(sheet_names):
    names = []
    for sheet_name in sheet_names:
        sharp_removed_name = sheet_name.replace('#', '')
        parenthesis_removed_name = sharp_removed_name.replace('(', '').replace(')', '').replace('（', '').replace('）', '',)

        match_result = re.match('.*?(\d+)', parenthesis_removed_name)
        if not match_result:
            code_name = '_CODE_'
            print('can not extract code name: {}'.format(parenthesis_removed_name))
        else:
            code_name = match_result.group()

        if '用户画像' in parenthesis_removed_name:
            module = '用户画像'
        elif '运营统计' in parenthesis_removed_name:
            module = '运营统计'
        elif '混合' in parenthesis_removed_name:
            module = '混合'
        else:
            module = '_MODULE_'
            print('can not extract module: {}'.format(parenthesis_removed_name))

        if 'android' in parenthesis_removed_name.lower():
            os = 'android'
        elif 'ios' in parenthesis_removed_name.lower():
            os = 'ios'
        elif module == '用户画像' or module == '混合':
            os = 'android'
        else:
            os = '_OS_'
            print('can not extract os: {}'.format(parenthesis_removed_name))

        real_name = parenthesis_removed_name.lower().replace(code_name, '').replace(module, '').replace(os, '').strip()
        if not real_name:
            print('no real name: {}'.format(parenthesis_removed_name))

        names.append(
            dict(old_name=sheet_name, code_name=code_name, module=module, os=os, real_name=real_name)
        )
    return names


def main():
    filename = common.get_filename()
    workbook = common.get_workbook(filename)
    names = normalize_sheet_names(common.get_sheet_names(workbook))
    old_name2new = {}
    for name in names:
        old_name2new.setdefault(name['old_name'], '{} {} {} {}'.format(name['code_name'], name['real_name'], name['module'], name['os']))
    for worksheet in common.list_worksheets(workbook):
        common.set_worksheet_name(worksheet, old_name2new[worksheet.title])
    common.save_workbook(workbook, filename)


if __name__ == '__main__':
    main()
