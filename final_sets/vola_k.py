import numpy as np
import scipy.stats as stats
import pylab as pl
import xlrd
import csv
import math
import random
import xlsxwriter
from sklearn.cluster import KMeans
import sys
from collections import namedtuple
"""
 y=y.reshape(1,-1)
>>> kmeans=KMeans(n_clusters=2, random_state=0).fit(y)
Traceback (most recent call last):

"""


workbook = xlrd.open_workbook("/home/ashwath/FinancialModel/final_sets/Yearly_Volatility.xlsx")
volatility = np.array([])
unfiltered = np.array([])
#def appendForYearI(sheetIndex,ebitda):
#for i in range(1,7):
sheet = workbook.sheet_by_name("VOLATILITY1")
for r in range(2,sheet.nrows):
    val=sheet.cell_value(r,1)
    #print val
    volatility=np.append(volatility,float(val))
    unfiltered=np.append(unfiltered,float(val))

#    print "Appended for Year ",i
mins=int(min(volatility))
maxs=int(max(volatility))
adds = int(((mins-maxs)/2)*0.5)
mins = mins - adds
maxs = maxs - adds
#for x in range(10000):
 #   volatility = np.append(volatility,random.uniform(min(volatility),max(volatility)))
  #  print(x)
length=volatility.size
print (len(volatility))
volatility=volatility.reshape(length,-1)

kmeans=KMeans(n_clusters=3, random_state=0).fit(volatility)
labels=kmeans.labels_
#print labels
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
#print l
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
#print ("this is", ds)
#print(int(max(volatility)))
#print(int(min(volatility)))
MyStruct = namedtuple("MyStruct", "cluster1 cluster2 cluster3")
m = MyStruct(centers[0], centers[1], centers[2])
#print(kmeans.cluster_centers_)
inc = 1;
sol = kmeans.predict(volatility)
#print(sol)
lim = 0
results = np.array([])
#print(len(unfiltered))
#for i in range(0,len(volatility)):
 #   print(volatility[i])
count = 0
for j in range(0, len(unfiltered)):
    count = 0
    for i in range(0,len(volatility)):
        if ((unfiltered[j])==(volatility[i]) and count==0):
            temp = sol[i]
            #print(sol[i])
            point = centers[temp]
            #print(point)
            dist = unfiltered[j] - point
            #print(volatility[i])
            results = np.append(results, float(dist))
            count = 1
            continue
print(len(volatility))
print(len(unfiltered))
print(len(results))

