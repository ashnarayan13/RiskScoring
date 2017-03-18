import numpy as np
import scipy.stats as stats
import pylab as pl
import xlrd
import csv
import math

def calcVar(ColIndex):
    workbook = xlrd.open_workbook("ins_modified.xlsx")
    sheet = workbook.sheet_by_name("TS")
    index=int(ColIndex);
    current = []
    x=0
    for rnum in range(3,sheet.nrows-1):
        x=math.log(sheet.cell(rnum+1,index).value/sheet.cell(rnum,index).value)
        current.append(x)
    print len(current)
    """
    for rows in range(3, sheet.nrows - 1):
        y = sheet.cell_value(rows, 5)
        z = sheet.cell_value(rows - 1, 5)
        current.append(math.log(y / z, math.e))
    """
    print current
   # fit = stats.norm.pdf(current, np.mean(current), np.std(current))  # this is a fitting indeed
    #pl.plot(current, fit, '-o')
    #pl.hist(current, bins=50, normed=True)
    print np.percentile(current, q=5)
   # pl.show()

ColIndex=raw_input("Enter Column No :")
calcVar(ColIndex)
