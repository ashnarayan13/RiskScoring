import xlrd
import numpy
import xlwt
def volatility(inp):
    print(inp)
    workbook = xlrd.open_workbook("ins_modified.xlsx")
    sheet = workbook.sheet_by_name("TS")
    allinfo = []
    const = inp*2
    rows = 4
    tot = 0
    cn = 0
    cn1 = 0
    ctr = 0
    retm = []
    y0 = 3
    y1 = 239
    y2 = 493
    y3 = 746
    y4 = 999
    y5 = 1253
    y6 = 1510
    y7 = 1766
    y8 = 2020
    y9 = 2274
    y10 = 2526
    y11 = 2781
    yearList = [y0,y1,y2,y3,y4,y5,y6,y7,y8,y9,y10,y11]
    for i in range(0,11):
        current = []
        for rows in range(yearList[i],yearList[i+1]-1):
            if(sheet.cell_value(rows,const)==' '):
                print("Missing value")
            else:
                #current = []
                current.append(float(sheet.cell_value(rows,const)))
                allinfo.append(current)
                temp_str = sheet.cell_value(rows,const)
                allinfo.append(float(temp_str))
                tot = tot + float(temp_str)
                vals = numpy.log(float(sheet.cell_value(rows,const))/float(sheet.cell_value(rows+1,const)))
                retm.append(float(vals))

#average returns
        ravg = []
        temp = 0
        for j in range(len(retm)):
            if(ctr!=21):
                temp = temp + retm[j]
                ctr = ctr + 1
            if(ctr>=21):
               ravg.append(float(temp/21))
               ctr = 0

#squared dev 
        ravg_t = 0
        chk = 0
        partial = 0.0
        temp_std = []
        for j in range(len(retm)):
            if((j+1)%22 == 0):
                temp_std.append(partial)
                ravg_t = ravg_t + 1
                partial = 0
            else:
                partial = partial + (retm[j]-ravg[ravg_t])**2
        std_val = sum(temp_std)/20
        print("Daily volatility : {}".format(std_val))
        std_val_annual = std_val*(252**0.5)
        print("Annualized Volatility : {}".format(std_val_annual))
        std_val_month = std_val*(12**0.5)
        print("Monthly Volatility : {}".format(std_val_month))
        #wb_sam_sheet.write(inp,1,std_val)
        #wb_sam_sheet.write(inp,2,std_val_month)
        if(i==0):
            wb_sam_sheet1.write(inp,1,std_val_annual)
        elif(i == 1):
            wb_sam_sheet2.write(inp,1,std_val_annual)
        elif(i == 2):
            wb_sam_sheet3.write(inp,1,std_val_annual)
        elif(i == 3):
            wb_sam_sheet4.write(inp,1,std_val_annual)
        elif(i == 4):
            wb_sam_sheet5.write(inp,1,std_val_annual)
        elif(i == 5):
            wb_sam_sheet6.write(inp,1,std_val_annual)
        elif(i == 6):
            wb_sam_sheet7.write(inp,1,std_val_annual)
        elif(i == 7):
            wb_sam_sheet8.write(inp,1,std_val_annual)
        elif(i == 8):
            wb_sam_sheet9.write(inp,1,std_val_annual)
        elif(i == 9):
            wb_sam_sheet10.write(inp,1,std_val_annual)
        elif(i == 10):
            wb_sam_sheet11.write(inp,1,std_val_annual)
#print list of available companies

company = "instruments.xlsm"
workbook = xlrd.open_workbook(company)
comp_sheet = workbook.sheet_by_name("CDAX")
companies = []
for j in range(2,comp_sheet.nrows):
    print(str(j-1)+" "+ str(comp_sheet.cell_value(j,3)))
    companies.append(str(comp_sheet.cell_value(j,3)))
#worksheet = int(raw_input())
#overall_vals = int(len(comp_sheet.nrows))
#print(overall_vals)
#print(companies[worksheet])
print("writing start")
wb_sam = xlwt.Workbook()
wb_sam_sheet1 = wb_sam.add_sheet("VOLATILITY1")
wb_sam_sheet2 = wb_sam.add_sheet("VOLATILITY2")
wb_sam_sheet3 = wb_sam.add_sheet("VOLATILITY3")
wb_sam_sheet4 = wb_sam.add_sheet("VOLATILITY4")
wb_sam_sheet5 = wb_sam.add_sheet("VOLATILITY5")
wb_sam_sheet6 = wb_sam.add_sheet("VOLATILITY6")
wb_sam_sheet7 = wb_sam.add_sheet("VOLATILITY7")
wb_sam_sheet8 = wb_sam.add_sheet("VOLATILITY8")
wb_sam_sheet9 = wb_sam.add_sheet("VOLATILITY9")
wb_sam_sheet10 = wb_sam.add_sheet("VOLATILITY10")
wb_sam_sheet11 = wb_sam.add_sheet("VOLATILITY11")
wb_sam_sheet1.write(0,3,"Annual Volatility")
wb_sam_sheet2.write(0,3,"Annual Volatility")
wb_sam_sheet3.write(0,3,"Annual Volatility")
wb_sam_sheet4.write(0,3,"Annual Volatility")
wb_sam_sheet5.write(0,3,"Annual Volatility")
wb_sam_sheet6.write(0,3,"Annual Volatility")
wb_sam_sheet7.write(0,3,"Annual Volatility")
wb_sam_sheet8.write(0,3,"Annual Volatility")
wb_sam_sheet9.write(0,3,"Annual Volatility")
wb_sam_sheet10.write(0,3,"Annual Volatility")
wb_sam_sheet11.write(0,3,"Annual Volatility")
for j in range(2, comp_sheet.nrows):
    print(comp_sheet.cell_value(j,3))
    wb_sam_sheet1.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet2.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet3.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet4.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet5.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet6.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet7.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet8.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet9.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet10.write(j,0,str(comp_sheet.cell_value(j,3)))
    wb_sam_sheet11.write(j,0,str(comp_sheet.cell_value(j,3)))
    volatility(j)
wb_sam.save("vol123.xlsx")
#volatility(worksheet)
