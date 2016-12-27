import numpy as np
import scipy.stats as stats
import pylab as pl
import xlrd
import csv
import math

def calcVar(fileName):
    workbook = xlrd.open_workbook(fileName)
    sheet = workbook.sheet_by_name("Tabelle1")
    current = []
    for rows in range(3, sheet.nrows - 1):
        y = sheet.cell_value(rows, 5)
        z = sheet.cell_value(rows - 1, 5)
        current.append(math.log(y / z, math.e))
    print current
    fit = stats.norm.pdf(current, np.mean(current), np.std(current))  # this is a fitting indeed
    pl.plot(current, fit, '-o')
    pl.hist(current, normed=True)
    print np.percentile(current, q=5)
    pl.show()

fileName=raw_input("Enter file name :")
calcVar(fileName)


