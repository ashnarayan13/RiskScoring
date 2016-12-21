'''
Created on Dec 21, 2016

@author: rajat
'''
import xlrd
from numpy import cov
from numpy import var
from matplotlib.testing.jpl_units import day


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
            print(indexSheet.cell_value(rows,4))
            print(equitySheeet.cell_value(rows+1,5))
            #date = indexSheet.cell_value(rows,0)
            #if indexSheet.cell_value(rows,0) == 3: # 3 means 'xldate' , 1 means 'text'
            #ms_date_number = indexSheet.cell_value(rows,0) # Correct option 1
                #ms_date_number = sheet.cell(5, 19).value # Correct option 2
            #year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ms_date_number,indexWorkBook.datemode)
                #py_date = datetime.datetime(year, month, day, hour, minute, nearest_second)
            #print(year+" "+month+" "+day)
            self.indexClosingPrice.append(indexSheet.cell_value(rows,4))
            self.equityClosingPrice.append(equitySheeet.cell_value(rows+1,5))
            
        print("Printing daily return of index")
        print(self.indexClosingPrice[3])
        
        
   
    def getBeta(self):
       #calcualte return for equity and index
       equityDailyReturn = []
       indexyDailyReturn = []
       print("In beta")
       for i in range(len(self.indexClosingPrice)):
            equityDailyReturn.append(((self.equityClosingPrice[i]-self.equityClosingPrice[i-1])/self.equityClosingPrice[i-1])*100)
            indexyDailyReturn.append(((self.indexClosingPrice[i]-self.indexClosingPrice[i-1])/self.indexClosingPrice[i-1])*100)
       print("Index return : ",indexyDailyReturn)
       print("Equity return : ",equityDailyReturn)  
       beta = cov(equityDailyReturn,indexyDailyReturn)[0][1]/var(indexyDailyReturn)
       print("covariance : ",cov(equityDailyReturn,indexyDailyReturn)[0][1])
       print("variance : ",var(indexyDailyReturn))
        
       return beta   
       
          
        
        
