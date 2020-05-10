import xlrd

try:
    workbook = xlrd.open_workbook('p30_till_36.xlsx')
except:
    print "error: file open"
sheet = workbook.sheet_by_index(0)
t = sheet.cell(0, 0).value
x = sheet.cell(0, 1).value
y = sheet.cell(0, 2).value
z = sheet.cell(0, 3).value
i = "Index"
#print "[ "+str(i)+" ]\t[ "+str(t)+" (ms) ]\t[ "+str(x)+" ]\t[ "+str(y)+" ]\t[ "+str(z)+" ]"
print "[ %-6s ] [ %-9s (ms) ] [ %-10s ] [ %-10s ] [ %10s ]"%(i,t,x,y,z)
#row, col

#total rows in the sheet
nrows = sheet.nrows

#total cols in the sheet
ncols = sheet.ncols

t = 0
for i in range(1, nrows):
    try:
        #t = round(float(str(sheet.cell(i, 0).value)), 4)
        x = round(float(str(sheet.cell(i, 1).value)), 4)
        y = round(float(str(sheet.cell(i, 2).value)), 4)
        z = round(float(str(sheet.cell(i, 3).value)), 4)
        
        while abs(x) > 4.0:
            x = x/10

        while abs(y) > 4.0:
            y = y/10

        while abs(z) > 4.0:
            z = z/10

        x = round(x, 4)
        y = round(y, 4)
        z = round(z, 4)


        print "[ %6s ] [ %14s ] [ %11s ] [ %11s ] [ %11s ]"%(i,t,x,y,z)
        #print "[ "+str(i)+" ]\t[ "+str(t)+" ]\t[ "+str(x)+" ]\t[ "+str(y)+" ]\t[ "+str(z)+" ]"
        t = t + 33.33;
        t = round(t, 4);
    except:
        pass
        '''
        t = "ERR"
        x = "ERR"
        y = "ERR"
        z = "ERR"
        print "[ "+str(i)+" ]\t[ "+str(t)+" ]\t[ "+str(x)+" ]\t[ "+str(y)+" ]\t[ "+str(z)+" ]"
        '''



