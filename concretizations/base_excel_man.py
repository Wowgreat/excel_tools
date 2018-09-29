import os

class BaseExcelMan():
    '''
    Base Excle man class, all Excel man must inherit from this class
    '''
    lines_len = 0
    columns_len = 0
    workbook = None
    worksheet = None
    char_list = {}
    titles = {}
    file_path  = None

    def __init__(self, file_path, sheet_name=None, has_title=True):
        self.file_path = file_path
        self.char_list['A-Z'] = [chr(i) for i in range(65, 91)]
        self.workbook, self.worksheet = \
            self.initialization(file_path=file_path, sheet_name=sheet_name)
        if os.path.exists(file_path):
            self.columns_len = self.get_lines_and_columns_num(1)
            self.lines_len = self.get_lines_and_columns_num('A')
            if has_title:
                self.get_title()

    def initialization(self, file_path, sheet_name):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def get_title(self):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def get_lines_and_columns_num(self, line_or_col):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def get_lines_and_columns(self, line_or_col):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def write_row(self, row_value_list, istitle=False):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def write_spicified_row(self, specified_row, row_value_list):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def write_specified_grid(self, value, line, title=None, col_name=None):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def del_specified_rows(self, specified_row, del_blank_rows=False):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def del_specified_cols(self, specified_cols, del_blank_cols=False):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))
