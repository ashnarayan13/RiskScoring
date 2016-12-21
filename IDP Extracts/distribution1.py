import numpy as np
import scipy.stats as stats
import pylab as pl
import xlrd

h = sorted([186, 176, 158, 180, 186, 168, 168, 164, 178, 170, 189, 195, 172,
     187, 180, 186, 185, 168, 179, 178, 183, 179, 170, 175, 186, 159,
     161, 178, 175, 185, 175, 162, 173, 172, 177, 175, 172, 177, 180])  #sorted
#workbook = xlrd.open_workbook("Adidas_2006_Now_D.xlsx")
workbook=xlrd.open_workbook("DAX_2006_Now_Index_D.xlsx")
sheet = workbook.sheet_by_name("Tabelle1")
for rows in range(2, 2782):
    current = []
    current.append(sheet.cell_value(rows,7))
    print current
fit = stats.norm.pdf(current, np.mean(current), np.std(current))  #this is a fitting indeed

pl.plot(current,fit,'-o')
print np.size(current)
pl.hist(current,normed=True)      #use this to draw histogram of your data

pl.show()


