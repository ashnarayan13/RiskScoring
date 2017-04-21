import numpy as np
import scipy.stats as stats
import pylab as pl
import xlrd
import xlwt
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
book = xlwt.Workbook()

#def appendForYearI(sheetIndex,ebitda):
#for i in range(1,7):
sheets = ["VOLATILITY1","VOLATILITY2","VOLATILITY3","VOLATILITY4","VOLATILITY5","VOLATILITY6","VOLATILITY7","VOLATILITY8","VOLATILITY9","VOLATILITY10","VOLATILITY11"]
for vols in range(0,len(sheets)):
    print("reading sheet {}".format(vols))
    sheet = workbook.sheet_by_name(sheets[vols])
    sheetwrite = book.add_sheet(sheets[vols])
    sheetwrite.write(0,1,"label")
    sheetwrite.write(0,2,"cluster center")
    sheetwrite.write(0,3,"dist")
    for r in range(3,sheet.nrows):
        val=sheet.cell_value(r,1)
        if(val == 'NULL'):
            continue
        #print val
        volatility=np.append(volatility,float(val))
        unfiltered=np.append(unfiltered,float(val))
        sheetwrite.write(r,0,str(sheet.cell_value(r,0)))
    mins=int(min(volatility))
    maxs=int(max(volatility))
    adds = int(((mins-maxs)/2)*0.5)
    mins = mins - adds
    maxs = maxs - adds
    for x in range(1000):
        volatility = np.append(volatility,random.uniform(min(volatility),max(volatility)))
        #print(x)
    length=volatility.size
    #print ("length of", length)
    volatility=volatility.reshape(length,-1)
    kmeans=KMeans(n_clusters=3, random_state=0).fit(volatility)
    labels=kmeans.labels_
    #print labels
    centers = kmeans.cluster_centers_
    #print ("This is the k mean centers",centers)
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
    #pl.scatter(dz,yy, c=None, s=500)
    #pl.show()
    MyStruct = namedtuple("MyStruct", "cluster1 cluster2 cluster3")
    m = MyStruct(centers[0], centers[1], centers[2])
#print(kmeans.cluster_centers_)
    sol = kmeans.predict(volatility)
    results = np.array([])
    lim = 3
    for j in range(0, len(unfiltered)):
        for i in range(0,len(volatility)):
            if ((unfiltered[j])==(volatility[i])):
                temp = sol[i]
                sheetwrite.write(lim,1,int(temp))
                #print(sol[i])
                point = centers[temp]
                sheetwrite.write(lim,2,float(point))
                #print(point)
                dist = unfiltered[j] - point
                sheetwrite.write(lim,3,float(dist))
                #print(dist)
                results = np.append(results, float(dist))
                lim = lim + 1
                continue
    #with open('somefile.txt', 'a') as the_file:
     #   the_file.write('Centers for {0} is {1} \n'.format(vols, centers))
book.save("kmeansResults.xlsx")
