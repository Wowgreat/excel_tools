from concretizations.excel_man_openpyxl import ExcelManOpenPyXl
import time
t1 = time.time()
excel = ExcelManOpenPyXl('../test.xlsx', has_title=True)

print(excel.titles)
excel.del_specified_cols( titles=['链接', '是否扫描'], del_blank_cols=False)
excel.write_specified_grid(value=500, line=5, title='是否为板块')
excel.del_specified_rows([7,8], del_blank_rows=True)
print(time.time()-t1)