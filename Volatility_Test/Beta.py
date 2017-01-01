'''
Created on Dec 21, 2016

@author: rajat
'''
import xlrd
from numpy import cov
from numpy import var
from matplotlib.testing.jpl_units import day
from datetime import datetime


class BetaWithTimeHoraizon:
    '''
    create beta(relative volatility ) for a specified time horizon 
    '''


    def __init__(self, startDate,endDate,equitySheeet):
        '''
        Constructor
        '''
        #Parsing index
        indexWorkBook = xlrd.open_workbook("DAXIndex.xlsx") 
        indexSheet = indexWorkBook.sheet_by_name("Sheet1")
                
        self.startDate = startDate
        self.endDate = endDate
        self.indexClosingPrice = []
        self.equityClosingPrice = []
        
        
        for rows in range(1,indexSheet.nrows):
            dateValue = datetime.strptime(indexSheet.cell_value(rows,0),'%Y-%m-%d').date()
            #get data between time horizoion 
            if dateValue<=startDate and dateValue>=endDate :
                self.indexClosingPrice.append(indexSheet.cell_value(rows,4))
            if endDate>dateValue :
                break
                   
            #if indexSheet.cell_value(rows,0) == 3: # 3 means 'xldate' , 1 means 'text'
            #ms_date_number = indexSheet.cell_value(rows,0) # Correct option 1
                #ms_date_number = sheet.cell(5, 19).value # Correct option 2
            #year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ms_date_number,indexWorkBook.datemode)
                #py_date = datetime.datetime(year, month, day, hour, minute, nearest_second)
            #print(year+" "+month+" "+day)
        for rows in range(2,equitySheeet.nrows):
            dateValue = datetime(*xlrd.xldate_as_tuple(equitySheeet.cell_value(rows,5), 0)).date()            
            if dateValue<=startDate and dateValue>=endDate :
                #taking closing price
                self.equityClosingPrice.append(equitySheeet.cell_value(rows,2))
            if endDate>dateValue :
                break
            
                 
   
    def getBeta(self):
       #calcualte return for equity and index
       equityDailyReturn = []
       indexyDailyReturn = []
       for i in range(len(self.indexClosingPrice)-1):                   
            equityDailyReturn.append(((self.equityClosingPrice[i+1]-self.equityClosingPrice[i])/self.equityClosingPrice[i])*100)
            indexyDailyReturn.append(((self.indexClosingPrice[i+1]-self.indexClosingPrice[i])/self.indexClosingPrice[i])*100)
       #print("Index return : ",indexyDailyReturn)
       #print("Equity return : ",equityDailyReturn)  
       #print("Equity closing price : ",self.equityClosingPrice)
       #print("Index closing price : ",self.indexClosingPrice)
       beta = cov(equityDailyReturn,indexyDailyReturn)[0][1]/var(indexyDailyReturn)
       #print("covariance : ",cov(equityDailyReturn,indexyDailyReturn)[0][1])
       #print("variance : ",var(indexyDailyReturn))
       #print("length of index size : ",len(indexyDailyReturn))
       #print("length of equity size : ",len(equityDailyReturn))
        
       return beta   
       
          
        
        