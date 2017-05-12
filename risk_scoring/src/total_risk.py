import numpy as np
import xlrd
import xlwt


def read_risks(sheet):
    risk = np.array([])
    for i in range(2, sheet.nrows):
        temp = float(sheet.cell_value(i, 4))
        risk = np.append(risk, temp)
    return risk


def risk_calculator(temp_ebitda, ebitda_dup, temp_volatility, volatility_dup, temp_roce, roce_dup, temp_gear, gear_dup,
                    temp_beta, beta_dup, temp_var, var_dup):
    temp = [ebitda_dup, volatility_dup, roce_dup, gear_dup, beta_dup, var_dup]
    fact = sum(temp)
    for i in range(0, len(temp)):
        temp[i] = fact - temp[i]
    total = float(sum(temp))
    print(total)
    risks = [(temp[0]) / total, (temp[1]) / total, (temp[2]) / total, (temp[3]) / total, (temp[4]) / total,
             (temp[5]) / total]
    risk_vals = np.zeros(425)
    print(sum(risks))
    for pts in range(0, len(temp_beta)):
        risk_temp = temp_ebitda[pts] * risks[0] + temp_volatility[pts] * risks[1] + temp_roce[pts] * risks[2] + \
                    temp_gear[pts] * risks[3] + temp_beta[pts] * risks[4] + temp_var[pts] * risks[5]
        risk_vals[pts] = risk_temp
    print(max(risk_vals))
    return risk_vals


book = xlwt.Workbook()
ebitda = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/risk_results/PyCharmEBITDA.xlsx")
volatility = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/risk_results/PyCharmVOLATILITY.xlsx")
roce = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/risk_results/PyCharmROCE.xlsx")
gear = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/risk_results/PyCharmGEAR.xlsx")
beta = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/risk_results/PyCharmBETA.xlsx")
var = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/risk_results/PyCharmVAR.xlsx")
writesheets = ["FY-1", "FY-2", "FY-3", "FY-4", "FY-5", "FY-6"]
tags = ["EBITDA", "ROCE", "VOLATILITY", "GEAR"]
final_ebitda = np.zeros(425)
final_volatility = np.zeros(425)
final_roce = np.zeros(425)
final_gear = np.zeros(425)
final_beta = np.zeros(425)
final_var = np.zeros(425)
final_IDP = np.zeros(425)
for i in range(0, len(writesheets)):
    temp_ebitda = read_risks(ebitda.sheet_by_name(writesheets[i]))
    ebitda_dup = int(ebitda.sheet_by_name(writesheets[i]).cell_value(0, 5))
    temp_volatility = read_risks(volatility.sheet_by_name(writesheets[i]))
    volatility_dup = int(volatility.sheet_by_name(writesheets[i]).cell_value(0, 5))
    temp_roce = read_risks(roce.sheet_by_name(writesheets[i]))
    roce_dup = int(roce.sheet_by_name(writesheets[i]).cell_value(0, 5))
    temp_gear = read_risks(gear.sheet_by_name(writesheets[i]))
    gear_dup = int(gear.sheet_by_name(writesheets[i]).cell_value(0, 5))
    temp_beta = read_risks(beta.sheet_by_name(writesheets[i]))
    beta_dup = int(beta.sheet_by_name(writesheets[i]).cell_value(0, 5))
    temp_var = read_risks(var.sheet_by_name(writesheets[i]))
    var_dup = int(var.sheet_by_name(writesheets[i]).cell_value(0, 5))
    total_risk = risk_calculator(temp_ebitda, ebitda_dup, temp_volatility, volatility_dup, temp_roce, roce_dup,
                                 temp_gear, gear_dup, temp_beta, beta_dup, temp_var, var_dup)
    for k in range(0, len(total_risk)):
        final_IDP[k] += total_risk[k]/6
print(min(final_IDP))
print(max(final_IDP))
sheet = book.add_sheet("results")
sheet.write(0, 0, "COMPANY")
sheet.write(0, 1, "RISK RATING")
results = ebitda.sheet_by_name(writesheets[i])
for vals in range(0, len(final_IDP)):
    sheet.write(vals + 2, 0, str(results.cell_value(vals + 2, 0)))
    sheet.write(vals + 2, 1, final_IDP[vals])
book.save("/home/ashwath/PycharmProjects/risk_scoring/risk_results/RESULTS.xlsx")
