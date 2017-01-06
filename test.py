import xlrd
import numpy

def return_stock(inp,loc):
    print(inp)
    print(loc)
    workbook = xlrd.open_workbook("worksheet" + str(loc+1) + ".xlsx")
    sheet = workbook.sheet_by_name("Tabelle1")
    store = []
    for rows in range(2, sheet.nrows-1):
        temp = numpy.log(float(sheet.cell_value(rows,2))/float(sheet.cell_value(rows+1,2)))
        store.append(float(temp))




#print list of available companies
company = "worksheet1.xlsx"
workbook = xlrd.open_workbook(company)
comp_sheet = workbook.sheet_by_name("Tabelle3")
companies = []
for i in range(1,comp_sheet.nrows):
    print(str(i-1)+" "+ str(comp_sheet.cell_value(i,0)))
    companies.append(str(comp_sheet.cell_value(i,0)))
worksheet = int(raw_input())
#print(companies[worksheet])
return_stock(companies[worksheet], worksheet)
