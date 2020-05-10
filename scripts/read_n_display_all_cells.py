import xlrd
file_name = 'p15.xlsx'
workbook = xlrd.open_workbook(file_name)
sheet = workbook.sheet_by_index(0)
#total rows in the sheet
nrows = sheet.nrows

#total cols in the sheet
ncols = sheet.ncols

for r in range(0, nrows):
    print("")
    for c in range(0, ncols):
        try:
            cell_value = str(sheet.cell(r, c).value)
            print("[ "+cell_value+" ]\t", end = '')
        except:
            pass



