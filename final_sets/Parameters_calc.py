import xlrd
import matplotlib.pyplot as plt
import xlwt
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
            up = up + 1
    print("EBITDA ",up)
    plt.subplot(321)
    plt.title("EBITDA")
    plt.plot(info)
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
            up = up + 1
    print("ROE ",up)
    plt.subplot(322)
    plt.title("ROE")
    plt.plot(roe)
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
            up = up + 1
    print("GEARING ",up)
    plt.subplot(323)
    plt.title("GEAR")
    plt.plot(gear)
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
            up = up + 1
    print("PROFITABILITY ",up)
    plt.subplot(324)
    plt.title("PROFITABILITY")
    plt.plot(prof)
#ROCE works! Using the equation EBIT/(ASSETS-LIABILITY)
def ROCE(choice, sheet41, sheet42, sheet43):
    roce = []
    for i in range(2, sheet41.ncols):
        if(sheet41.cell_value(choice, i) != "NULL" and sheet42.cell_value(choice, i) != "NULL" and sheet43.cell_value(choice, i) != "NULL"):
            roce.append(float(sheet41.cell_value(choice, i))/(float(sheet42.cell_value(choice, i))-float(sheet43.cell_value(choice, i))))
    up = 0
    for i in range(1, len(roce)):
        if(roce[i] >= roce[i-1]):
            up = up+1
    print("ROCE ",up)
    plt.subplot(325)
    plt.plot(roce)
    plt.title("ROCE")
    plt.subplots_adjust(top=0.92, bottom=0.08, left=0.10, right=0.95, hspace=0.25,
                    wspace=0.35)
    plt.show()

#############################################
workbook1 = xlrd.open_workbook("/home/ashwath/FinancialModel/final_sets/Data_Collection_Code/EBITDA.xlsx")
sheet11 = workbook1.sheet_by_name("EBITDA")
sheet12 = workbook1.sheet_by_name("SALES")
for rows in range(2, sheet11.nrows):
    print(rows-1, str(sheet11.cell_value(rows,1)))
choice = int(raw_input())
choice = choice + 1
EBIT(choice, sheet11)
PROFITABILITY(choice, sheet11, sheet12)

workbook2 = xlrd.open_workbook("/home/ashwath/FinancialModel/final_sets/Data_Collection_Code/RETURN_EQUITY.xlsx")
sheet21 = workbook2.sheet_by_name("TOTAL_REVENUE")
sheet22 = workbook2.sheet_by_name("EQUITY")
sheet23 = workbook2.sheet_by_name("COGS")
ROE(choice, sheet21, sheet22, sheet23)

workbook3 = xlrd.open_workbook("/home/ashwath/FinancialModel/final_sets/Data_Collection_Code/GEARING.xlsx")
sheet31 = workbook3.sheet_by_name("DEBT")
sheet32 = workbook3.sheet_by_name("EQUITY")
GEAR(choice, sheet31, sheet32)

workbook4 = xlrd.open_workbook("/home/ashwath/FinancialModel/final_sets/Data_Collection_Code/ROCE.xlsx")
sheet41 = workbook4.sheet_by_name("EBIT")
sheet42 = workbook4.sheet_by_name("ASSETS")
sheet43 = workbook4.sheet_by_name("LIABILITY")
ROCE(choice, sheet41, sheet42, sheet43)
