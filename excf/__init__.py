import os
import xlrd


class VcfCreator(object):
    def __init__(self, sh, name_col_num=0, phone_col_num=1, filename=None, vcard_name=None):
        """
        :param sh: `sheet class object` the sheet include name and phone
        :param name_col_num: column number of names
        :param phone_col_num: column number of phone numbers
        :param filename: name of file you want to open
        :param vcard_name: name of vcard you want to create
        """
        os.getcwd().encode('utf-8')  # set coding to utf-8
        self.info_list = []
        self.read_file = ''
        self.name_col_number = name_col_num
        self.phone_col_number = phone_col_num
        self.filename = filename
        self.vcard_name = vcard_name

    def read_excel(self, filename=None, is_title_exist=True):
        """
        read the file info and store it into the object
        """
        filename = filename or self.filename
        assert (isinstance(is_title_exist, bool))
        if filename:
            try:
                # Use slice to remove the title
                with xlrd.open_workbook(filename) as f:
                    sh = f.sheet_by_index(0)
                    name_col = sh.col_values(self.name_col_number)[int(is_title_exist):]
                    phone_col = sh.col_values(self.phone_col_number)[int(is_title_exist):]
                    self.info_list = zip(name_col, phone_col)
                    self.read_file = filename
            except NameError as why:
                raise why
        else:
            raise IOError('You should input filename')

    def create_vcf(self, vcard_name='contact.vcf'):
        vcard_name = vcard_name or self.vcard_name
        if vcard_name:
            try:
                # Using slice to remove the title
                with open(vcard_name, 'w', encoding='utf-8') as f:
                    for name, phone in self.info_list:
                        f.write('BEGIN:VCARD\n')
                        f.write(f'FN:{name}\n')
                        f.write(f'TEL;type=CELL;type=VOICE;type=pref:{phone}\n')
                        f.write('VERSION:3.0\n')
                        f.write('END:VCARD\n')
            except (IOError, TypeError) as why:
                raise why