from concretizations.excel_man_openpyxl import ExcelManOpenPyXl

excel = ExcelManOpenPyXl('one.xlsx', has_title=True)
print(excel.get_lines_and_columns(title='title'))