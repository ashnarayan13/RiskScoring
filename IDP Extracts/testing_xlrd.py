import xlrd
import numpy

#reading from table
workbook = xlrd.open_workbook("Adidas_2006_Now_D.xlsx")
sheet = workbook.sheet_by_name("Tabelle1")
allinfo = []
rows = 3
tot = 0
cn = 0
cn1 = 0
ctr = 0
#print(sheet.nrows)
retm = []
for rows in range(2, sheet.nrows-1):
    current = []
    current.append(sheet.cell_value(rows,2))
    allinfo.append(current)
    temp_str = sheet.cell_value(rows,2)
    #store = float(temp_str)
    allinfo.append(float(temp_str))
    tot = tot + float(temp_str)
    #if(ctr!=21):
    vals = numpy.log(float(sheet.cell_value(rows,2))/float(sheet.cell_value(rows+1,2)))
    retm.append(float(vals))

#average returns
ravg = []
temp = 0
for i in range(len(retm)):
    if(ctr!=21):
        temp = temp + retm[i]
        ctr = ctr + 1
    if(ctr >= 21):
       ravg.append(float(temp/21))
       ctr = 0
#print(len(ravg))

#squared dev 
ravg_t = 0
chk = 0
partial = 0
temp_std = []
for i in range(len(retm)):
    if((i+1)%22 == 0):
        temp_std.append(partial)
        ravg_t = ravg_t + 1
        partial = 0
    partial = partial + (retm[i]-ravg[ravg_t])**2
std_val = sum(temp_std)/20

std_val = std_val*(252**0.5)
print(std_val)
#print(temp_str)
#print(tot)
#print(allinfo)
#print(return_month)
