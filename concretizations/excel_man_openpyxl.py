from openpyxl import load_workbook

from .base_excel_man import BaseExcelMan


class ExcelManOpenPyXl(BaseExcelMan):

    def initialization(self, file_path, sheet_name):
        wb = load_workbook(filename=file_path)
        if sheet_name is None:
            worksheet = wb[wb.sheetnames[0]]
        else:
            worksheet = wb[sheet_name]
        return worksheet

    def get_lines_and_columns(self, line_or_col=None, title=None):
        values = []
        if title is None:
            line_or_col = line_or_col
        else:
            line_or_col = self.titles[title]
        for val_obj in self.worksheet[line_or_col]:
            values.append(val_obj.value)
        return values

    def get_lines_and_columns_num(self,line_or_col):
        return len(self.get_lines_and_columns(line_or_col))

    def get_title(self):
        '''
        :return:
        '''
        title_line = self.get_lines_and_columns(1)
        for i in range(0, len(title_line)):
            k = title_line[i]
            if i <= 25:
                v = self.char_list['A-Z'][i]

            else:
                index = int(i / 26) - 1
                first_letter = self.char_list['A-Z'][index]
                second_letter = self.char_list['A-Z'][i - 26]
                v = first_letter + second_letter
            self.titles[k] = v
