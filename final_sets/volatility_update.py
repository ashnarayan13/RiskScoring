import xlrd
import numpy
def volatility(inp):
    print(inp)
    workbook = xlrd.open_workbook("ins_modified.xlsx")
    sheet = workbook.sheet_by_name("TS")
    allinfo = []
    const = inp*2
    #for rows in range(4, sheet.nrows-1):
     #   print(sheet.cell_value(rows,const))
    rows = 4
    tot = 0
    cn = 0
    cn1 = 0
    ctr = 0
#print(sheet.nrows)
    retm = []
    for rows in range(3, sheet.nrows-1):
        if(sheet.cell_value(rows,const)==' '):
            print("Missing value")
        else:
            #print(str(sheet.cell_value(rows,const))) debugging
            current = []
            current.append(float(sheet.cell_value(rows,const)))
            allinfo.append(current)
            temp_str = sheet.cell_value(rows,const)
    #store = float(temp_str)
            allinfo.append(float(temp_str))
            tot = tot + float(temp_str)
    #if(ctr!=21):
            vals = numpy.log(float(sheet.cell_value(rows,const))/float(sheet.cell_value(rows+1,const)))
            retm.append(float(vals))

#average returns
    ravg = []
    temp = 0
    for i in range(len(retm)):
        if(ctr!=21):
            temp = temp + retm[i]
            ctr = ctr + 1
        if(ctr>=21):
           ravg.append(float(temp/21))
           ctr = 0
#print(len(ravg))

#squared dev 
    ravg_t = 0
    chk = 0
    partial = 0.0
    temp_std = []
    for i in range(len(retm)):
        if((i+1)%22 == 0):
            temp_std.append(partial)
            ravg_t = ravg_t + 1
            partial = 0
        else:
            partial = partial + (retm[i]-ravg[ravg_t])**2
    std_val = sum(temp_std)/20
    #print(worksheet) debugging
    print("Daily volatility : {}".format(std_val))
    std_val_annual = std_val*(252**0.5)
    print("Annualized Volatility : {}".format(std_val_annual))
    std_val_month = std_val*(12**0.5)
    print("Monthly Volatility : {}".format(std_val_month))
#print list of available companies
company = "instruments.xlsm"
workbook = xlrd.open_workbook(company)
comp_sheet = workbook.sheet_by_name("CDAX")
companies = []
for i in range(2,comp_sheet.nrows):
    print(str(i-1)+" "+ str(comp_sheet.cell_value(i,3)))
    companies.append(str(comp_sheet.cell_value(i,3)))
worksheet = int(raw_input())
#print(companies[worksheet])
volatility(worksheet)
