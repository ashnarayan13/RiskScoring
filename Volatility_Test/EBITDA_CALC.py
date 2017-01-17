import xlwt
import xlrd

def EBIT(choice, sheet):
    info = []
    for cols in range(2, sheet.ncols):
        info.append(int(sheet.cell_value(choice,cols)))
    #print info
    up = 0
    for i in range(0, len(info)-1):
        if(info[i]<=info[i+1]):
            #print(info[i], " lesser ", info[i+1])
            up = up + 1
    print("EBITDA ",up)

def ROE(choice, sheet21, sheet22):
    roe = []
    for i in range(2, sheet21.ncols):
        roe.append(float((float(sheet21.cell_value(choice, i))/float(sheet22.cell_value(choice, i)))*100))
    #print(roe)
    up = 0
    for i in range(0, len(roe)-1):
        if(roe[i]<=roe[i+1]):
            #print(roe[i], " lesser ", roe[i+1])
            up = up + 1
    print("ROE ",up)

def GEAR(choice, sheet31, sheet32):
    gear= []
    for i in range(2, sheet31.ncols):
        gear.append(float((float(sheet31.cell_value(choice, i))/float(sheet32.cell_value(choice, i)))*100))
    #print(gear)
    up = 0
    for i in range(0, len(gear)-1):
        if(gear[i]<=gear[i+1]):
            #print(info[i], " lesser ", info[i+1])
            up = up + 1
    print("GEARING ",up)

def PROFITABILITY(choice, sheet11, sheet12):
    prof= []
    for i in range(2, sheet11.ncols):
        if(sheet11.cell_value(choice, i) != "NULL" and sheet12.cell_value(choice, i) != "NULL"):
            prof.append(float((float(sheet11.cell_value(choice, i))/float(sheet12.cell_value(choice, i)))*100))
    #print(prof)
    up = 0
    for i in range(0, len(prof)-1):
        if(prof[i]<=prof[i+1]):
            #print(info[i], " lesser ", info[i+1])
            up = up + 1
    print("PROFITABILITY ",up)

#############################################
workbook1 = xlrd.open_workbook("EBITDA.xlsx")
sheet11 = workbook1.sheet_by_name("EBITDA")
sheet12 = workbook1.sheet_by_name("SALES")
for rows in range(2, sheet11.nrows):
    print(rows-1, str(sheet11.cell_value(rows,1)))
choice = int(raw_input())
choice = choice + 1
EBIT(choice, sheet11)
PROFITABILITY(choice, sheet11, sheet12)

workbook2 = xlrd.open_workbook("RETURN_EQUITY.xlsx")
sheet21 = workbook2.sheet_by_name("REV_AFTER_TAX")
sheet22 = workbook2.sheet_by_name("EQUITY")
ROE(choice, sheet21, sheet22)

workbook3 = xlrd.open_workbook("GEARING.xlsx")
sheet31 = workbook3.sheet_by_name("DEBT")
sheet32 = workbook3.sheet_by_name("EQUITY")
GEAR(choice, sheet31, sheet32)

