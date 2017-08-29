import risk
import xlrd
import xlwt

print("List of parameters available: \n 1) Volatility \n 2) Value at Risk \n 3) Return on Equity \n 4) Gearing \n 5) Profitability \n 6) Beta \n 7) Return on Capital Employed \n 8) EBITDA"
      "\n Return on Equity and Profitabiliy does not work due to poor data quality")

choice = int(input('Enter the parameter to be calculated: '))
if choice == 1:
    file_name = "VOLATILITY"
elif choice == 2:
    file_name = "VAR"
elif choice == 6:
    file_name = "BETA"
else:
    file_name = "FY_Params"
read_file = "/home/ashwath/FinancialModel/risk_scoring/Data_Collection_Code/" + file_name + ".xlsx"
workbook = xlrd.open_workbook(read_file)
book = xlwt.Workbook()

writesheets = ["FY-1", "FY-2", "FY-3", "FY-4", "FY-5", "FY-6"]

sheetNames = ["Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5", "Sheet6"]
volatility_sheets = ["VOLATILITY1", "VOLATILITY2", "VOLATILITY3", "VOLATILITY4", "VOLATILITY5", "VOLATILITY6"]

for lims in range(0, len(writesheets)):
    # Mention column to read from along with the sheet
    if choice == 1:
        readsheet = workbook.sheet_by_name(volatility_sheets[lims])
        values, used = risk.create_data(readsheet, 1)
        setting = 1
        save_file = "VOLATILITY"
    elif choice == 2:
        readsheet = workbook.sheet_by_name(sheetNames[lims])
        values, used = risk.create_data(readsheet, 2)
        setting = 1
        save_file = "VAR"
    elif choice == 3:
        readsheet = workbook.sheet_by_name("ROE")
        values, used = risk.create_data(readsheet, lims+2)
        setting = 2
        save_file = "ROE"
    elif choice == 5:
        readsheet = workbook.sheet_by_name("PROFITABILITY")
        values, used = risk.create_data(readsheet, lims+2)
        setting = 2
        save_file = "PROFITABILITY"
    elif choice == 8:
        readsheet = workbook.sheet_by_name("EBITDA")
        values, used = risk.create_data(readsheet, lims+2)
        setting = 2
        save_file = "EBITDA"
    elif choice == 7:
        readsheet = workbook.sheet_by_name("ROCE")
        values, used = risk.create_data(readsheet, lims+2)
        setting = 2
        save_file = "ROCE"
    elif choice == 4:
        readsheet = workbook.sheet_by_name("GEAR")
        values, used = risk.create_data(readsheet, lims+2)
        setting = 2
        save_file = "GEAR"
    elif choice == 6:
        readsheet = workbook.sheet_by_name("sheet1")
        values, used = risk.create_data(readsheet, lims+1)
        setting = 3
        save_file = "BETA"

    # Calculte the duplicates
    duplicates = risk.detect_duplicates(values)
    # setup kmeans network
    km, updated_used = risk.kmeans_setup(used)
    # to get the mapping, distance from centroid, sort centers
    maps, dists, ranges, sorted_centers = risk.map_distance(km, updated_used, values)
    print("printing ranges")
    print(ranges)
    print("printing sorted range")
    print(sorted_centers)
    print("mapping")
    print(maps)
    # get the risk values
    answers = risk.risk_vals(ranges, values, sorted_centers, maps, setting)
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
if choice==1 or choice==2 or choice==6 or choice==8 or choice==4 or choice==7:
    save_file = "/home/ashwath/FinancialModel/risk_scoring/risk_results/risk_" + save_file + ".xlsx"
    book.save(save_file)