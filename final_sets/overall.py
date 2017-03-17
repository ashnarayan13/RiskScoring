import xlrd
import numpy
import matplotlib.pyplot as plt
import xlwt

####################################
#VOLATILITY CALCULATION
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
    for rows in range(3, sheet.nrows-1):
        if(sheet.cell_value(rows,const)==' '):
            print("Missing value")
        else:
            current = []
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
    for i in range(len(retm)):
        if(ctr!=21):
            temp = temp + retm[i]
            ctr = ctr + 1
        if(ctr>=21):
           ravg.append(float(temp/21))
           ctr = 0

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
    print("Daily volatility : {}".format(std_val))
    std_val_annual = std_val*(252**0.5)
    print("Annualized Volatility : {}".format(std_val_annual))
    std_val_month = std_val*(12**0.5)
    print("Monthly Volatility : {}".format(std_val_month))
    wb_sam_sheet.write(inp,1,std_val)
    wb_sam_sheet.write(inp,2,std_val_month)
    wb_sam_sheet.write(inp,3,std_val_annual)
#print list of available companies

company = "instruments.xlsm"
workbook = xlrd.open_workbook(company)
comp_sheet = workbook.sheet_by_name("CDAX")
companies = []
for i in range(2,comp_sheet.nrows):
    print(str(i-1)+" "+ str(comp_sheet.cell_value(i,3)))
    companies.append(str(comp_sheet.cell_value(i,3)))
#worksheet = int(raw_input())
#overall_vals = int(len(comp_sheet.nrows))
#print(overall_vals)
#print(companies[worksheet])
wb_sam = xlwt.Workbook()
wb_sam_sheet = wb_sam.add_sheet("VOLATILITY")
wb_sam_sheet.write(0,1,"Daily Volatility")
wb_sam_sheet.write(0,2,"Monthly Volatility")
wb_sam_sheet.write(0,3,"Annual Volatility")
for i in range(2, comp_sheet.nrows):
    print(comp_sheet.cell_value(i,3))
    wb_sam_sheet.write(i,0,str(comp_sheet.cell_value(i,3)))
    volatility(i)
wb_sam.save("PARAMETERS.xlsx")

##############################################################
#EBITDA CALC
#EBIT Calculation is correct! Compares the EBITDA values of the years
def EBIT(choice, sheet):
    info = []
    for cols in range(2, sheet.ncols):
        if(sheet.cell_value(choice, cols) != "NULL"):
            info.append(int(sheet.cell_value(choice,cols)))
    #print info
    up = 0
    for i in range(1, len(info)):
        if(info[i]>=info[i-1]):
            #print(info[i], " lesser ", info[i+1])
            wb_EBITDA.write(choice,i,str(sheet11.cell_value(rows,1)))
            up = up + 1
    print("EBITDA ",up)
    #plt.subplot(321)
    #plt.title("EBITDA")
    #plt.plot(info)
#ROE works! Follows the equation taken from the lecture (Total_Revenue - Cost_Of_Goods_Sold)/EQUITY
def ROE(choice, sheet21, sheet22, sheet23):
    roe = []
    for i in range(2, sheet21.ncols):
        if(sheet21.cell_value(choice, i) != "NULL" and sheet22.cell_value(choice, i) != "NULL" and sheet23.cell_value(choice, i) != "NULL"):
            roe.append(float((float(sheet21.cell_value(choice, i)-sheet23.cell_value(choice, i))/float(sheet22.cell_value(choice, i)))*100))
    #print(roe)
    up = 0
    for i in range(1, len(roe)):
        if(roe[i]>=roe[i-1]):
            #print(roe[i], " lesser ", roe[i+1])
            wb_ROE.write(choice,i,str(sheet11.cell_value(rows,1)))
            up = up + 1
    print("ROE ",up)
    #plt.subplot(322)
    #plt.title("ROE")
    #plt.plot(roe)
#GEAR works! will output 0 if the debt is 0 as it can't be used otherwise using the equation debt/equity
def GEAR(choice, sheet31, sheet32):
    gear= []
    for i in range(2, sheet31.ncols):
        if(sheet31.cell_value(choice, i) != "NULL" and sheet32.cell_value(choice, i) != "NULL"):
            if(sheet31.cell_value(choice, i) != 0):
                gear.append(float((float(sheet31.cell_value(choice, i))/float(sheet32.cell_value(choice, i)))*100))
    #print(gear)
    up = 0
    for i in range(1, len(gear)):
        if(gear[i]<=gear[i-1]):
            #print(info[i], " lesser ", info[i+1])
            wb_GEAR.write(choice,i,str(sheet11.cell_value(rows,1)))
            up = up + 1
    print("GEARING ",up)
    #plt.subplot(323)
    #plt.title("GEAR")
    #plt.plot(gear)
#Profitability works! Using the equation EBITDA/Net_sale 
def PROFITABILITY(choice, sheet11, sheet12):
    prof= []
    for i in range(2, sheet11.ncols):
        if(sheet11.cell_value(choice, i) != "NULL" and sheet12.cell_value(choice, i) != "NULL"):
            prof.append(float((float(sheet11.cell_value(choice, i))/float(sheet12.cell_value(choice, i)))*100))
    #print(prof)
    up = 0
    for i in range(1, len(prof)):
        if(prof[i]>=prof[i-1]):
            #print(info[i], " lesser ", info[i+1])
            wb_PROF.write(choice,i,str(sheet11.cell_value(rows,1)))
            up = up + 1
    print("PROFITABILITY ",up)
    #plt.subplot(324)
    #plt.title("PROFITABILITY")
    #plt.plot(prof)
#ROCE works! Using the equation EBIT/(ASSETS-LIABILITY)
def ROCE(choice, sheet41, sheet42, sheet43):
    roce = []
    for i in range(2, sheet41.ncols):
        if(sheet41.cell_value(choice, i) != "NULL" and sheet42.cell_value(choice, i) != "NULL" and sheet43.cell_value(choice, i) != "NULL"):
            roce.append(float(sheet41.cell_value(choice, i))/(float(sheet42.cell_value(choice, i))-float(sheet43.cell_value(choice, i))))
    up = 0
    for i in range(1, len(roce)):
        if(roce[i] >= roce[i-1]):
            wb_ROCE.write(choice,i,str(sheet11.cell_value(rows,1)))
            up = up+1
    print("ROCE ",up)
    #plt.subplot(325)
    #plt.plot(roce)
    #plt.title("ROCE")
    #plt.subplots_adjust(top=0.92, bottom=0.08, left=0.10, right=0.95, hspace=0.25,
     #               wspace=0.35)
    #plt.show()
'''
#############################################
workbook1 = xlrd.open_workbook("EBITDA.xlsx")
sheet11 = workbook1.sheet_by_name("EBITDA")
sheet12 = workbook1.sheet_by_name("SALES")
wb_EBITDA = wb_sam.add_sheet("EBITDA")
wb_PROF = wb_sam.add_sheet("PROFITABILITY")

workbook2 = xlrd.open_workbook("RETURN_EQUITY.xlsx")
sheet21 = workbook2.sheet_by_name("TOTAL_REVENUE")
sheet22 = workbook2.sheet_by_name("EQUITY")
sheet23 = workbook2.sheet_by_name("COGS")
wb_ROE = wb_sam.add_sheet("ROE")

workbook3 = xlrd.open_workbook("GEARING.xlsx")
sheet31 = workbook3.sheet_by_name("DEBT")
sheet32 = workbook3.sheet_by_name("EQUITY")
wb_GEAR = wb_sam.add_sheet("GEAR")

workbook4 = xlrd.open_workbook("ROCE.xlsx")
sheet41 = workbook4.sheet_by_name("EBIT")
sheet42 = workbook4.sheet_by_name("ASSETS")
sheet43 = workbook4.sheet_by_name("LIABILITY")
wb_ROCE = wb_sam.add_sheet("ROCE")

for rows in range(2, sheet11.nrows):
    print(rows-1, str(sheet11.cell_value(rows,1)))
    EBIT(rows, sheet11)
    wb_EBITDA.write(rows,0,str(sheet11.cell_value(rows,1)))
    PROFITABILITY(rows, sheet11, sheet12)
    wb_PROF.write(rows,0,str(sheet11.cell_value(rows,1)))
    ROE(rows, sheet21, sheet22, sheet23)
    wb_ROE.write(rows,0,str(sheet11.cell_value(rows,1)))
    GEAR(rows, sheet31, sheet32)
    wb_GEAR.write(rows,0,str(sheet11.cell_value(rows,1)))
    ROCE(rows, sheet41, sheet42, sheet43)
    wb_ROCE.write(rows,0,str(sheet11.cell_value(rows,1)))

wb_sam.save("PARAMETERS.xlsx")



    
choice = int(raw_input())
choice = choice + 1
#EBIT(choice, sheet11)
#PROFITABILITY(choice, sheet11, sheet12)

workbook2 = xlrd.open_workbook("RETURN_EQUITY.xlsx")
sheet21 = workbook2.sheet_by_name("TOTAL_REVENUE")
sheet22 = workbook2.sheet_by_name("EQUITY")
sheet23 = workbook2.sheet_by_name("COGS")
ROE(choice, sheet21, sheet22, sheet23)

workbook3 = xlrd.open_workbook("GEARING.xlsx")
sheet31 = workbook3.sheet_by_name("DEBT")
sheet32 = workbook3.sheet_by_name("EQUITY")
GEAR(choice, sheet31, sheet32)

workbook4 = xlrd.open_workbook("ROCE.xlsx")
sheet41 = workbook4.sheet_by_name("EBIT")
sheet42 = workbook4.sheet_by_name("ASSETS")
sheet43 = workbook4.sheet_by_name("LIABILITY")
ROCE(choice, sheet41, sheet42, sheet43)
'''
