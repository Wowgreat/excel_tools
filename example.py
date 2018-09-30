from excel import Excel

excel = Excel('test.xlsx', has_title=True).instance

print(excel.titles)