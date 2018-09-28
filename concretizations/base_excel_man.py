import os

class BaseExcelMan():
    '''
    Base Excle man class, all Excel man must inherit from this class
    '''
    lines_len = 0
    columns_len = 0
    worksheet = None
    char_list = {}
    titles = {}
    def __init__(self, file_path, sheet_name=None, has_title=True):
        if os.path.exists(file_path):
            self.worksheet = self.initialization(file_path=file_path, sheet_name=sheet_name)
            self.columns_len = self.get_lines_and_columns_num(1)
            self.lines_len = self.get_lines_and_columns_num('A')
            if has_title:
                self.char_list['A-Z'] = [chr(i) for i in range(65, 91)]
                self.get_title()

    def initialization(self, file_path, sheet_name):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def get_title(self):
        '''
        obtain a dict like as {'title1': 'A'}, and assign the result to the attribute titles
        '''
        pass

    def get_lines_and_columns_num(self, line_or_col):
        '''
        obtain the length of lines or columns and return
        '''
        pass

    def get_lines_and_columns(self, line_or_col):
        '''
        obtain the lines or columns of excel and return
        '''
        pass

    def write_one_row(self):
        pass

    def write_spicified_row(self, specified_row, row_value):
        pass

    def write_specified_grid(self, specified_grid, value):
        pass

    def del_specified_row(self,specified_row):
        pass
