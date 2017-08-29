import fy_parameters
import risk_scoring_fy_params
import xlrd
import xlwt


'''Make datasets with parameters'''
wb_sam = xlwt.Workbook()
workbook1 = xlrd.open_workbook("/home/ashwath/FinancialModel/risk_scoring/Data_Collection_Code/EBITDA.xlsx")
sheet11 = workbook1.sheet_by_name("EBITDA")
sheet12 = workbook1.sheet_by_name("SALES")
wb_EBITDA = wb_sam.add_sheet("EBITDA")
wb_PROF = wb_sam.add_sheet("PROFITABILITY")

workbook2 = xlrd.open_workbook("/home/ashwath/FinancialModel/risk_scoring/Data_Collection_Code/RETURN_EQUITY.xlsx")
sheet21 = workbook2.sheet_by_name("TOTAL_REVENUE")
sheet22 = workbook2.sheet_by_name("EQUITY")
sheet23 = workbook2.sheet_by_name("COGS")
wb_ROE = wb_sam.add_sheet("ROE")

workbook3 = xlrd.open_workbook("/home/ashwath/FinancialModel/risk_scoring/Data_Collection_Code/GEARING.xlsx")
sheet31 = workbook3.sheet_by_name("DEBT")
sheet32 = workbook3.sheet_by_name("EQUITY")
wb_GEAR = wb_sam.add_sheet("GEAR")

workbook4 = xlrd.open_workbook("/home/ashwath/FinancialModel/risk_scoring/Data_Collection_Code/ROCE.xlsx")
sheet41 = workbook4.sheet_by_name("EBIT")
sheet42 = workbook4.sheet_by_name("ASSETS")
sheet43 = workbook4.sheet_by_name("LIABILITY")
wb_ROCE = wb_sam.add_sheet("ROCE")

for rows in range(2, sheet11.nrows):
    print(rows - 1, str(sheet11.cell_value(rows, 1)))
    fy_parameters.ebit(rows, sheet11, wb_EBITDA)
    wb_EBITDA.write(rows, 0, str(sheet11.cell_value(rows, 1)))
    fy_parameters.profitability(rows, sheet11, sheet12, wb_PROF)
    wb_PROF.write(rows, 0, str(sheet11.cell_value(rows, 1)))
    fy_parameters.roe(rows, sheet21, sheet22, sheet23, wb_ROE)
    wb_ROE.write(rows, 0, str(sheet11.cell_value(rows, 1)))
    fy_parameters.gear(rows, sheet31, sheet32, wb_GEAR)
    wb_GEAR.write(rows, 0, str(sheet11.cell_value(rows, 1)))
    fy_parameters.roce(rows, sheet41, sheet42, sheet43, wb_ROCE)
    wb_ROCE.write(rows, 0, str(sheet11.cell_value(rows, 1)))

wb_sam.save("/home/ashwath/FinancialModel/risk_scoring/Data_Collection_Code/FY_Params.xlsx")


'''Calculate risk scores'''

workbook = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/Data_Collection_Code/FY_Params.xlsx")
book = xlwt.Workbook()

writesheets = ["FY-1", "FY-2", "FY-3", "FY-4", "FY-5", "FY-6"]

writesheets_1 = ["VOLATILITY1", "VOLATILITY2", "VOLATILITY3", "VOLATILITY4", "VOLATILITY5", "VOLATILITY6"]
#readsheet = workbook.sheet_by_name("sheet1")
for lims in range(0, len(writesheets)):
    # Mention column to read from along with the sheet
    readsheet = workbook.sheet_by_name("GEAR")
    values, used = risk_scoring_fy_params.create_data(readsheet, lims + 2)
    # Calculte the duplicates
    duplicates = risk_scoring_fy_params.detect_duplicates(values)
    # setup kmeans network
    km, used = risk_scoring_fy_params.kmeans_setup(used)
    # to get the mapping, distance from centroid, sort centers
    maps, dists, ranges, sorted_centers = risk_scoring_fy_params.map_distance(km, used, values)
    print("printing ranges")
    print(ranges)
    print("printing sorted range")
    print(sorted_centers)
    print("mapping")
    print(maps)
    # get the risk values
    answers = risk_scoring_fy_params.risk_vals(ranges, values, sorted_centers, maps)
    sheetwrite = book.add_sheet(writesheets[lims])
    sheetwrite.write(0, 0, "COMPANY")
    sheetwrite.write(0, 1, "VALUE")
    sheetwrite.write(0, 2, "CLUSTER CENTER")
    sheetwrite.write(0, 3, "DISTANCE")
    sheetwrite.write(0, 4, "RISK RATING")
    sheetwrite.write(0, 5, duplicates)
    for k in range(0, len(values)):
        sheetwrite.write(k + 2, 0, str(readsheet.cell_value(k + 2, 0)))
        sheetwrite.write(k + 2, 1, values[k])
        sheetwrite.write(k + 2, 3, dists[k])
        sheetwrite.write(k + 2, 4, answers[k])
book.save("/home/ashwath/PycharmProjects/risk_scoring/risk_results/PyCharmGEAR.xlsx")