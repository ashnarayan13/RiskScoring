import numpy as np
import scipy.stats as stats
import pylab as pl
import xlrd
import csv
import math
import xlsxwriter

book=xlsxwriter.Workbook("VAR.xlsx")
wsheet=book.add_worksheet("VAR")

yr=2006
for i in range(0,11):
    wsheet.write(i,0,yr)
    yr=yr+1

col=1
def calcVar(ColIndex):

    index=int(ColIndex);
    current = []
    y0 = 3
    y1 = 239
    y2 = 493
    y3 = 746
    y4 = 999
    y5 = 1253
    y6 = 1510
    y7 = 1766
    y8 = 2020
    y9 = 2274
    y10 = 2526
    y11 = 2781
    yearList = [y0,y1,y2,y3,y4,y5,y6,y7,y8,y9,y10,y11]
    for i in range(0,11):
        current = []
        for rnum in range(yearList[i],yearList[i+1]-1):
            x=math.log(sheet.cell(rnum,index).value/sheet.cell(rnum+1,index).value)
            current.append(x)

        print len(current)
        print current
        """
        fit = stats.norm.pdf(current, np.mean(current), np.std(current))  # this is a fitting indeed
        pl.plot(current, fit, '-o')
        pl.hist(current, bins=50, normed=True)
        """
        var = np.percentile(current, q=5)
        print var
        wsheet.write(i,col,var)
        #pl.show()

workbook = xlrd.open_workbook("DataSet_Final.xlsx")
sheet = workbook.sheet_by_name("TS")
for i in range(3,sheet.ncols,2):
    #ColIndex=raw_input("Enter Column No :")
    calcVar(i)
    col=col+1;
    print "Done for Company ",i
