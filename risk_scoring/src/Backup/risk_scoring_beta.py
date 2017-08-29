# coding=utf-8
import numpy as np
import xlrd
import xlwt
import random
from sklearn.cluster import KMeans
import scipy


def create_data(sheet, col):
    values_cr_data = np.array([])
    used_cr_data = np.array([])
    for i in range(2, sheet.nrows):
        temp = float(readsheet.cell_value(i, col))
        values_cr_data = np.append(values_cr_data, temp)
        used_cr_data = np.append(used_cr_data, temp)
    return values_cr_data, used_cr_data


def detect_duplicates(sheet_read):
    temp_ctr = 0
    for el in set(sheet_read):
        if sheet_read.tolist().count(el) > 1:
            temp_ctr += sheet_read.tolist().count(el)
            # print el, sheet_read.count(el), len(sheet_read) - x[::-1].index(el)
    print(temp_ctr)
    return temp_ctr

def risk_vals(data):
    risk = np.array([])
    for i in data:
        if i <= 0.25 and i >= 0:
            risk = np.append(risk, 1)
        elif i <= 0.5 and i >= 0.25:
            risk = np.append(risk, 2)
        elif i <= 0.75 and i >= 0.5:
            risk = np.append(risk, 3)
        elif i <= 1 and i >= 0.75:
            risk = np.append(risk, 4)
        else:
            risk = np.append(risk, 5)
    return risk
# END OF FUNCTIONS

workbook = xlrd.open_workbook("/home/ashwath/PycharmProjects/risk_scoring/Data_Collection_Code/BetaYearWise.xlsx")
book = xlwt.Workbook()

writesheets = ["FY-1", "FY-2", "FY-3", "FY-4", "FY-5", "FY-6"]

writesheets_1 = ["VOLATILITY1", "VOLATILITY2", "VOLATILITY3", "VOLATILITY4", "VOLATILITY5", "VOLATILITY6"]
readsheet = workbook.sheet_by_name("sheet1")
for lims in range(0, len(writesheets)):
    # Mention column to read from along with the sheet
    values, used = create_data(readsheet, lims+1)
    # Calculte the duplicates
    duplicates = detect_duplicates(values)
    # get the risk values
    answers = risk_vals(values)
    sheetwrite = book.add_sheet(writesheets[lims])
    sheetwrite.write(0, 0, "COMPANY")
    sheetwrite.write(0, 1, "VALUE")
    sheetwrite.write(0, 4, "RISK RATING")
    sheetwrite.write(0, 5, duplicates)
    for k in range(0, len(values)):
        sheetwrite.write(k + 2, 0, str(readsheet.cell_value(k + 2, 0)))
        sheetwrite.write(k + 2, 4, answers[k])
book.save("/home/ashwath/PycharmProjects/risk_scoring/risk_results/PyCharmBETA.xlsx")
