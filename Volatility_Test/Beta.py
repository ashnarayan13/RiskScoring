'''
Created on Dec 21, 2016

@author: rajat
'''
from numpy import cov
from numpy import var

class BetaWithTimeHoraizon:
    '''
    create beta(relative volatility ) for a specified time horizon 
    '''


    def __init__(self, eqPrice,indexValue):
        '''
        Constructor
        '''
        self.equityClosingPrice = eqPrice
        self.indexClosingPrice = indexValue
        
   
      def getBeta(self):
       #calcualte return for equity and index
       equityDailyReturn = []
       indexDailyReturn = []
       for i in len(self)
         equityDailyReturn[i] = ((self.equityClosingPrice[i]-self.equityClosingPrice[i-1])/self.equityClosingPrice[i-1])*100
         indexyDailyReturn[i] = ((self.indexClosingPrice[i]-self.indexClosingPrice[i-1])/self.indexClosingPrice[i-1])*100
         
       beta = cov(equityDailyReturn,indexDailyReturn)/var(indexDailyReturn)
         
    return beta   
       
                
        
        