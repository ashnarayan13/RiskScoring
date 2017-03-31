import numpy as np
import scipy.stats as stats
import pylab as pl
import xlrd
import csv
import math
import random
import xlsxwriter
from sklearn.cluster import KMeans
"""
 y=y.reshape(1,-1)
>>> kmeans=KMeans(n_clusters=2, random_state=0).fit(y)
Traceback (most recent call last):

"""


workbook = xlrd.open_workbook("/home/ashwath/FinancialModel/final_sets/PARAMETERS.xlsx")
volatility = np.array([])
#def appendForYearI(sheetIndex,ebitda):
#for i in range(1,7):
sheet = workbook.sheet_by_name("VOLATILITY")
for r in range(3,sheet.nrows):
    val=sheet.cell_value(r,3)
    if(val == 'NULL'):
        continue
    print val
    volatility=np.append(volatility,int(val))

#    print "Appended for Year ",i
mins=min(volatility)
maxs=max(volatility)
adds = int(((mins-maxs)/2)*0.5)
mins = mins - adds
maxs = maxs - adds
for x in range(100000):
    volatility = np.append(volatility,random.randint(min(volatility),max(volatility)))
    print(x)
length=volatility.size
print ("length of", length)
volatility=volatility.reshape(length,-1)

kmeans=KMeans(n_clusters=3, random_state=0).fit(volatility)
labels=kmeans.labels_
print labels
centers = kmeans.cluster_centers_
print ("This is the k mean centers",centers)
ds = volatility[np.where(labels==0)]
dy = volatility[np.where(labels==1)]
dz = volatility[np.where(labels==2)]
l=len(ds)
yy = np.array([])
for i in range(0,l):
    yy=np.append(yy,1)
pl.scatter(ds,yy,c=None,s=500)
print l
ly=len(dy)
yy = np.array([])
for i in range(0,ly):
    yy=np.append(yy,2)
pl.scatter(dy,yy, c=None, s=500)

lz=len(dz)
yy = np.array([])
for i in range(0,lz):
    yy=np.append(yy,3)
pl.scatter(dz,yy, c=None, s=500)
pl.show()
print ("this is", ds)
print(int(max(volatility)))
print(int(min(volatility)))

