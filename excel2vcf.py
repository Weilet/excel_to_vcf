#! -*- coding : utf-8 -*-

import os
import xlrd


def read_excel(filename):
    """

    :param filename: 要打开的文件的路径
    :return success: excel表第一个sheet
            fail: 失败原因
    """
    try:
        with xlrd.open_workbook(filename) as f:
            return f.sheet_by_index(0)
    except IOError as e:
            return e


class VcfCreater(object):
    def __init__(self, sh):
        """

        :param sh: 包含姓名以及联系方式的sheet
        """
        o_path = os.getcwd().encode('utf-8')  #  将 coding 设置为 utf-8
        try:
            name_col = sh.col_values(0)[1:]
            phone_col = sh.col_values(1)[1:]
            with open('contact.vcf', 'w', encoding='utf-8') as f:
                for name, phone in zip(name_col, phone_col):
                    f.write('BEGIN:VCARD\n')
                    f.write(f'FN:{name}\n')
                    f.write(f'TEL;type=CELL;type=VOICE;type=pref:{phone}\n')
                    f.write('VERSION:3.0\n')
                    f.write('END:VCARD\n')
        except (IOError, TypeError) as e:
            print(e)


if __name__ == '__main__':
    try:
        filename = input('输入 excel 文件路径\n')
        sh = read_excel(filename)
        vcard = VcfCreater(sh)
    except IOError as e:
        print(e)
        exit(code=4)
