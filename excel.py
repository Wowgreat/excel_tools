class Excel():
    '''
    NO.1: Excel provide some common method for excel options,
    NO.2: Excel choice appropriate concretization Class base on parameters
    '''

    lines = 0
    columns = 0

    def __init__(self, filepath):
        pass

    # def __init__(self, filepath):
    #     wb = load_workbook(filename=filepath)
    #     self.work_sheet = wb[sheetname]
    #     if self.work_sheet is None:
    #         raise Exception
    #
    #     self.lines = len(self.work_sheet['A'])
    #     self.columns = len(self.work_sheet[1])
    #
    # def read_one_line_or_column(self, name):
    #     values = []
    #     for val_obj in self.work_sheet[name]:
    #         values.append(val_obj.value)
    #     return values
    #
    # def read_accurate_val(self, col_line):
    #     return self.work_sheet[col_line].value
