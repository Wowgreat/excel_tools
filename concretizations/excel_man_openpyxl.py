from openpyxl import load_workbook

from .base_excel_man import BaseExcelMan

class ExcelManOpenPyXl(BaseExcelMan):

    def parse(self, file_path, sheet_name):
        wb = load_workbook(filename=file_path)
        if sheet_name is None:
            worksheet = wb[wb.sheetnames[0]]
        else:
            worksheet = wb[sheet_name]
        return worksheet

    def get_lines_and_columns_num(self, line_or_col):
        values = []
        for val_obj in self.worksheet[line_or_col]:
            values.append(val_obj.value)
        return values

    def get_title(self):
        '''
        :return:
        '''
        title_line = self.get_lines_and_columns_num(1)
        for i in range(0, len(title_line)):
            if i <= 25:
                self.titles[title_line[i]] = self.char_list['A-Z'][i]
            elif i > 25:
                pass
