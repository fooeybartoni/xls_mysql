import xlrd

workbook = xlrd.open_workbook('/home/cnd_user/dev/53.xlsx')
worksheet = workbook.sheet_by_name('Rev4')
header_row = 3

print worksheet.ncols

headers = [worksheet.cell_value(header_row, i) for i in xrange(worksheet.ncols)]


row_dict_list = []
for row in xrange(header_row+1, worksheet.nrows):
    row_dict = {}
for col in xrange(worksheet.ncols):
    cell_type = worksheet.cell_type(row, col)
if cell_type == xlrd.XL_CELL_EMPTY:
    value = None
elif cell_type == xlrd.XL_CELL_TEXT:
    value = worksheet.cell_value(row, col)
elif cell_type == xlrd.XL_CELL_NUMBER:
    value = float(worksheet.cell_value(row,col))
elif cell_type == xlrd.XL_CELL_DATE:
    value = xlrd.xldate_as_tuple(worksheet.cell_value(row, col), workbook.datemode)
elif cell_type == xlrd.XL_CELL_BOOLEAN:
    value = bool(worksheet.cell_value(row, col))
else:
    value = worksheet.cell_value(row, col)
row_dict[headers[col]] = value
row_dict_list.append(row_dict)

for row in row_dict_list:
    print "*"*10
for col, value in row.iteritems():
    print col, value
    print "*"*10