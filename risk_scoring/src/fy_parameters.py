import xlrd
import xlwt


##############################################################
# EBITDA CALC
# EBIT Calculation is correct! Compares the EBITDA values of the years
def ebit(choice, sheet):
    info = []
    # print(sheet.ncols)
    for cols in range(2, 8):
        if sheet.cell_value(choice, cols) != "NULL":
            info.append(int(sheet.cell_value(choice, cols)))
            wb_EBITDA.write(choice, cols, int(sheet.cell_value(choice, cols)))
        else:
            info.append(0)
            wb_EBITDA.write(choice, cols, int(0))


# ROE works! Follows the equation taken from the lecture (Total_Revenue - Cost_Of_Goods_Sold)/EQUITY
def roe(choice, sheet21_roe, sheet22_roe, sheet23_roe):
    roe_cal = []
    for i in range(2, sheet21_roe.ncols):
        if sheet21_roe.cell_value(choice, i) != "NULL" and sheet22_roe.cell_value(choice, i) != "NULL" and sheet23_roe.cell_value(choice, i) != "NULL" and sheet22_roe.cell_value(choice, i) != 0:
            temp = float((float(sheet21_roe.cell_value(choice, i) - sheet23_roe.cell_value(choice, i)) / float(sheet22_roe.cell_value(choice, i))) * 100)
            roe_cal.append(float((float(sheet21_roe.cell_value(choice, i) - sheet23_roe.cell_value(choice, i)) / float(sheet22_roe.cell_value(choice, i))) * 100))
            wb_ROE.write(choice, i, temp)
        else:
            roe_cal.append(0)
            wb_ROE.write(choice, i, float(0))


# GEAR works! will output 0 if the debt is 0 as it can't be used otherwise using the equation debt/equity
def gear(choice, sheet31_gear, sheet32_gear):
    gear_cal = []
    for i in range(2, sheet31_gear.ncols):
        if sheet31_gear.cell_value(choice, i) != "NULL" and sheet32_gear.cell_value(choice, i) != "NULL":
            if sheet32_gear.cell_value(choice, i) != 0:
                temp = float((float(sheet31_gear.cell_value(choice, i)) / float(sheet32_gear.cell_value(choice, i))) * 100)
                gear_cal.append(float((float(sheet31_gear.cell_value(choice, i)) / float(sheet32_gear.cell_value(choice, i))) * 100))
                wb_GEAR.write(choice, i, temp)
            else:
                gear_cal.append(0)
                wb_GEAR.write(choice, i, float(0))
        else:
            gear_cal.append(0)
            wb_GEAR.write(choice, i, float(0))


# Profitability works! Using the equation EBITDA/Net_sale
def profitability(choice, sheet11_prof, sheet12_prof):
    prof = []
    for i in range(2, sheet11_prof.ncols):
        if sheet11_prof.cell_value(choice, i) != "NULL" and sheet12_prof.cell_value(choice, i) != "NULL" and sheet12_prof.cell_value(choice, i) != 0:
            temp = float((float(sheet11_prof.cell_value(choice, i)) / float(sheet12_prof.cell_value(choice, i))) * 100)
            prof.append(float((float(sheet11_prof.cell_value(choice, i)) / float(sheet12_prof.cell_value(choice, i))) * 100))
            wb_PROF.write(choice, i, temp)
        else:
            prof.append(0)
            wb_PROF.write(choice, i, float(0))
            # print(prof)


# ROCE works! Using the equation EBIT/(ASSETS-LIABILITY)
def roce(choice, sheet41_roce, sheet42_roce, sheet43_roce):
    roce_calc = []
    for i in range(2, sheet41_roce.ncols):
        if sheet41_roce.cell_value(choice, i) != "NULL" and sheet42_roce.cell_value(choice, i) != "NULL" and sheet43_roce.cell_value(choice, i) != "NULL":
            if sheet42_roce.cell_value(choice, i) != 0 and sheet43_roce.cell_value(choice, i) != 0:
                temp = float(sheet41_roce.cell_value(choice, i)) / (float(sheet42_roce.cell_value(choice, i)) - float(sheet43_roce.cell_value(choice, i)))
                roce_calc.append(float(sheet41_roce.cell_value(choice, i)) / (float(sheet42_roce.cell_value(choice, i)) - float(sheet43_roce.cell_value(choice, i))))
                wb_ROCE.write(choice, i, temp)
            else:
                roce_calc.append(0)
                wb_ROCE.write(choice, i, float(0))
        else:
            roce_calc.append(0)
            wb_ROCE.write(choice, i, float(0))


#############################################
wb_sam = xlwt.Workbook()
workbook1 = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/Data_Collection_Code/EBITDA.xlsx")
sheet11 = workbook1.sheet_by_name("EBITDA")
sheet12 = workbook1.sheet_by_name("SALES")
wb_EBITDA = wb_sam.add_sheet("EBITDA")
wb_PROF = wb_sam.add_sheet("PROFITABILITY")

workbook2 = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/Data_Collection_Code/RETURN_EQUITY.xlsx")
sheet21 = workbook2.sheet_by_name("TOTAL_REVENUE")
sheet22 = workbook2.sheet_by_name("EQUITY")
sheet23 = workbook2.sheet_by_name("COGS")
wb_ROE = wb_sam.add_sheet("ROE")

workbook3 = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/Data_Collection_Code/GEARING.xlsx")
sheet31 = workbook3.sheet_by_name("DEBT")
sheet32 = workbook3.sheet_by_name("EQUITY")
wb_GEAR = wb_sam.add_sheet("GEAR")

workbook4 = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/Data_Collection_Code/ROCE.xlsx")
sheet41 = workbook4.sheet_by_name("EBIT")
sheet42 = workbook4.sheet_by_name("ASSETS")
sheet43 = workbook4.sheet_by_name("LIABILITY")
wb_ROCE = wb_sam.add_sheet("ROCE")

for rows in range(2, sheet11.nrows):
    print(rows - 1, str(sheet11.cell_value(rows, 1)))
    ebit(rows, sheet11)
    wb_EBITDA.write(rows, 0, str(sheet11.cell_value(rows, 1)))
    profitability(rows, sheet11, sheet12)
    wb_PROF.write(rows, 0, str(sheet11.cell_value(rows, 1)))
    roe(rows, sheet21, sheet22, sheet23)
    wb_ROE.write(rows, 0, str(sheet11.cell_value(rows, 1)))
    gear(rows, sheet31, sheet32)
    wb_GEAR.write(rows, 0, str(sheet11.cell_value(rows, 1)))
    roce(rows, sheet41, sheet42, sheet43)
    wb_ROCE.write(rows, 0, str(sheet11.cell_value(rows, 1)))

wb_sam.save("/home/ashwath/PycharmProjects/risk_scoring/Data_Collection_Code/FY_Params.xlsx")
