class BaseExcelMan():
    '''
    Base Excle man class, all Excel man must inherit from this class
    '''
    lines = []
    columns = []
    lines_len = 0
    columns_len = 0
    worksheet = None
    char_list = {}
    has_title = None
    titles = {}
    # {
    #     format as 'A: title1' or 'AB: titlexx'
    # }

    def __init__(self, file_path, sheet_name=None, has_title=True):
        self.has_title = has_title
        self.worksheet = self.parse(file_path=file_path, sheet_name=sheet_name)
        if has_title:
            self.char_list['A-Z'] = [chr(i) for i in range(65, 91)]
            self.get_title()

    def parse(self, file_path, sheet_name):
        raise NotImplementedError('{}.parse callback is not defined'.format(self.__class__.__name__))

    def get_title(self):
        '''
        :return:
        '''
        pass

    def get_lines_and_columns_num(self, line_or_col):
        '''
        :return:
        '''
        pass

    def read_one_line_or_column(self):
        '''
        :return:
        '''
        pass

    def read_accurate_val(self):
        '''
        :return:
        '''
        pass
