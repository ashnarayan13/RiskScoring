import numpy as np
import scipy.stats as stats
import pylab as pl
import xlrd
import csv
import math
import xlsxwriter
import random
from sklearn.cluster import KMeans
"""
 y=y.reshape(1,-1)
>>> kmeans=KMeans(n_clusters=2, random_state=0).fit(y)
Traceback (most recent call last):

"""


workbook = xlrd.open_workbook("FY_modified.xlsx")
ebitda = np.array([])
#def appendForYearI(sheetIndex,ebitda):
for i in range(1,7):
    sheet = workbook.sheet_by_name("FY-"+str(i))
    for r in range(3,sheet.nrows-1):
        val=sheet.cell_value(r,3)
        if(val == 'NULL'):
            continue
        print val
        ebitda=np.append(ebitda,int(val))

    print "Appended for Year ",i


#for x in range(100000):
 #   ebitda = np.append(ebitda,random.randint(min(ebitda),max(ebitda)))
  #  print(x)
  
length=ebitda.size
print ("length of", length)

ebitda=ebitda.reshape(length,-1)

kmeans=KMeans(n_clusters=3, random_state=0).fit(ebitda)
labels=kmeans.labels_
print labels
centers = kmeans.cluster_centers_
print ("This is the k mean centers",centers)
ds = ebitda[np.where(labels==0)]
dy = ebitda[np.where(labels==1)]
dz = ebitda[np.where(labels==2)]
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
print(int(max(ebitda)))
print(int(min(ebitda)))

