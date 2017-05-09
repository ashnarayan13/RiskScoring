'''
Created on May 3, 2017

@author: Rajat
'''
import numpy as np
import scipy.stats as stats
import pylab as pl
import xlrd
import csv
import math
import xlsxwriter
from numpy import cov
from numpy import var
#book=xlsxwriter.Workbook("BetaYearWise.xlsx")
#wsheet=book.add_worksheet("sheet1")
# yr=2006
# for i in range(0,11):
#     wsheet.write(i,0,yr)
#     yr=yr+1
# 
# col=1
#global index value of daily index till 1 year
indexClosingValue = []

def getIndexClosingValue(rowNumber,sheet):
    sumOfClosingValue = 0
    #print('value of index at row ',rowNumber,' and total column no ',sheet.ncols)
    for i in range(3,sheet.ncols,2):
        sumOfClosingValue += sheet.cell(rowNumber,i).value
        #print(' ',sheet.cell(rowNumber,i).value)
    #return sum of all equity closing value  
    #print('Index of 1st row ',sumOfClosingValue);
    return sumOfClosingValue    
        

def calcBeta():
    book=xlsxwriter.Workbook("BetaYearWise.xlsx")
    wsheet=book.add_worksheet("sheet1")    
#initialize equity and index daily return 
    equityDailyReturn = []
    indexDailyReturn = []
    year=2006
    
    #index=int(ColIndex);
    #current = []
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
    for i in range(0,11):               #for 11 years of Data of daily open and close value            
                        
        currentYearBeta = []
        indexClosingValue = []
        indexDailyReturn = []
        sheet = workbook.sheet_by_name("TS")
        #print("No of row and col ",sheet.nrows,"  ",sheet.ncols)
        
        for eachCompany in range(3,sheet.ncols,2):
            equityDailyReturn = []
            for rowNum in range(yearList[i],yearList[i+1]-1):  #calculate 1 year equity and index row wise                
            #for rowNum in range(3,88):  #calculate 1 year equity and index row wise 
                #for all company index value will remain same 
                if(eachCompany == 3):
                    indexClosingValue.append(getIndexClosingValue(rowNum,sheet))
                    #print('index closing value : ',indexClosingValue,' and length ',len(indexClosingValue))
                if(rowNum > yearList[i]):
                    if (eachCompany == 3):
                       ## print("row and col number :",rowNum,"  ",eachCompany,' and index len ',len(indexClosingValue),'  and index pos ',(rowNum-1)-yearList[i])
                        indexDailyReturn.append(((indexClosingValue[rowNum-yearList[i]]- indexClosingValue[(rowNum-1)-yearList[i]])/indexClosingValue[(rowNum-1)-yearList[i]])*100)
                        #print('return value ',indexDailyReturn[rowNum-4],' value of last col ',sheet.cell(rowNum,3).value)
                 
                #for each company 
                    equityDailyReturn.append(((sheet.cell(rowNum,eachCompany).value - sheet.cell(rowNum-1,eachCompany).value)/sheet.cell(rowNum-1,eachCompany).value)*100)
                    #print('length of equityDailyReturn ',len(equityDailyReturn),' and index daily return ',len(indexDailyReturn))
                    #print('current day eq ',sheet.cell(rowNum,eachCompany).value,' after day equity ',sheet.cell(rowNum-1,eachCompany).value,'  and daily return ',equityDailyReturn[rowNum-4])
                #indexDailyReturn.append(((indexClosingValue[i+1]- indexClosingValue[i])/indexClosingValue[i])*100)
            #x=math.log(sheet.cell(rnum+1,).value/sheet.cell(rnum,index).value)
            #current.append(x)
            ##print('length of equityDailyReturn ',len(equityDailyReturn),' and index daily return ',len(indexDailyReturn),' for company ',((eachCompany+1)/2)-2)
            
            beta = cov(equityDailyReturn,indexDailyReturn)[0][1]/var(indexDailyReturn)
            print('value of beta ',beta,' for company ',((eachCompany+1)/2)-2,' for year ',i)                                   
            currentYearBeta.append(beta)
            wsheet.write(((eachCompany+1)/2)-2,i,beta)
            
        print('Current year beta ',currentYearBeta)
        
        """
        fit = stats.norm.pdf(current, np.mean(current), np.std(current))  # this is a fitting indeed
        pl.plot(current, fit, '-o')
        pl.hist(current, bins=50, normed=True)
        """
        #var = np.percentile(current, q=5)
        #print var
        #wsheet.write(i,col,np.transpose(currentYearBeta))
        #pl.show()
    print "Done for Company ",i    
        
 #open TS workbook for daily data       
workbook = xlrd.open_workbook("/home/rajat/workspace1/RiskScoreMatrix/FinancialModel/final_sets/TS_RAJAT.xlsx")
#sheet = workbook.sheet_by_name("TS")

#take closing value of each stock
#for i in range(3,sheet.ncols,2):
    #ColIndex=raw_input("Enter Column No :")
    
calcBeta()
    #col=col+1;
print('Program terminates ')    
