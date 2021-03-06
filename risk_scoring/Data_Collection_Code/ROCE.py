import xlrd
# 15 16 (ebit/assets-liabilities)
# or 8 19 (ebit/equity+debt)
from xlwt import Workbook

wb1 = xlrd.open_workbook("/home/ashwath/FinancialModel/final_sets/DataSet_Final.xlsx")
comp = wb1.sheet_by_name("CDAX")
s1 = wb1.sheet_by_name("FY-1")
s2 = wb1.sheet_by_name("FY-2")
s3 = wb1.sheet_by_name("FY-3")
s4 = wb1.sheet_by_name("FY-4")
s5 = wb1.sheet_by_name("FY-5")
s6 = wb1.sheet_by_name("FY-6")
wb2 = Workbook()
sheet1 = wb2.add_sheet("EBIT")
sheet2 = wb2.add_sheet("ASSETS")
sheet3 = wb2.add_sheet("LIABILITY")

for rows in range(2, comp.nrows):
    sheet1.write(rows, 0, comp.cell_value(rows, 2))
    sheet1.write(rows, 1, comp.cell_value(rows, 3))
    sheet2.write(rows, 0, comp.cell_value(rows, 2))
    sheet2.write(rows, 1, comp.cell_value(rows, 3))
    sheet3.write(rows, 0, comp.cell_value(rows, 2))
    sheet3.write(rows, 1, comp.cell_value(rows, 3))
for rows in range(3, s1.nrows):
    sheet1.write(rows - 1, 2, s1.cell_value(rows, 4))
    sheet1.write(rows - 1, 3, s2.cell_value(rows, 4))
    sheet1.write(rows - 1, 4, s3.cell_value(rows, 4))
    sheet1.write(rows - 1, 5, s4.cell_value(rows, 4))
    sheet1.write(rows - 1, 6, s5.cell_value(rows, 4))
    sheet1.write(rows - 1, 7, s6.cell_value(rows, 4))
    sheet2.write(rows - 1, 2, s1.cell_value(rows, 16))
    sheet2.write(rows - 1, 3, s2.cell_value(rows, 16))
    sheet2.write(rows - 1, 4, s3.cell_value(rows, 16))
    sheet2.write(rows - 1, 5, s4.cell_value(rows, 16))
    sheet2.write(rows - 1, 6, s5.cell_value(rows, 16))
    sheet2.write(rows - 1, 7, s6.cell_value(rows, 16))
    sheet3.write(rows - 1, 2, s1.cell_value(rows, 17))
    sheet3.write(rows - 1, 3, s2.cell_value(rows, 17))
    sheet3.write(rows - 1, 4, s3.cell_value(rows, 17))
    sheet3.write(rows - 1, 5, s4.cell_value(rows, 17))
    sheet3.write(rows - 1, 6, s5.cell_value(rows, 17))
    sheet3.write(rows - 1, 7, s6.cell_value(rows, 17))
wb2.save('ROCE.xlsx')
print("complete")
