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

workbook = xlrd.open_workbook("/home/ashwath/FinancialModel/final_sets/Yearly_Volatility.xlsx")
book = xlwt.Workbook()

#READ ALL THE SHEETS AND SOLVE MONTE CARLO FOR EACH SHEET 
sheets = ["VOLATILITY1","VOLATILITY2","VOLATILITY3","VOLATILITY4","VOLATILITY5","VOLATILITY6","VOLATILITY7","VOLATILITY8","VOLATILITY9","VOLATILITY10","VOLATILITY11"]
for vols in range(0,len(sheets)):
    volatility = np.array([])
    unfiltered = np.array([])
    print("reading sheet {}".format(vols))
    sheet = workbook.sheet_by_name(sheets[vols])
    sheetwrite = book.add_sheet(sheets[vols])
    sheetwrite.write(0,0,"COMPANY")
    sheetwrite.write(0,1,"LABEL")
    sheetwrite.write(0,2,"CLUSTER CENTER")
    sheetwrite.write(0,3,"DISTANCE")
    sheetwrite.write(0,4,"RISK RATING")
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
    #GENERATE RANDOM NUMBERS
    for x in range(10000):
        volatility = np.append(volatility,random.uniform(min(volatility),max(volatility)))
        #print(x)
    length=volatility.size
    #print ("length of", length)
    volatility=volatility.reshape(length,-1)
    #SOLVE KMEANS
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
    #PLOT POINTS IF NEEDED
    #pl.scatter(dz,yy, c=None, s=500)
    #pl.show()
    #COMPARE THE COMPANY RESULTS WITH CLUSTER CENTERS AND STORE THE DISTANCE
    mapping = kmeans.predict(volatility)
    results = np.array([])
    lim = 3
    maxval0 = float(centers[0])
    minval0 = float(centers[0])
    maxval1 = float(centers[1])
    minval1 = float(centers[1])
    maxval2 = float(centers[2])
    minval2 = float(centers[2])
    store = np.array([])
    for j in range(0, len(unfiltered)):
        count = 0
        for i in range(0,len(volatility)):
            if ((unfiltered[j])==(volatility[i]) and count==0):
                temp = mapping[i]
                store = np.append(store,temp)
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
                count = 1
                #TO GET THE RANGE FOR EACH CENTROID
                if temp == 0:
                    if unfiltered[j] >= maxval0:
                        maxval0 = unfiltered[j]
                    if unfiltered[j] <= minval0:
                        minval0 = unfiltered[j]
                if temp == 1:
                    if unfiltered[j] >= maxval0:
                        maxval1 = unfiltered[j]
                    if unfiltered[j] <= minval0:
                        minval1 = unfiltered[j]
                if temp == 2:
                    if unfiltered[j] >= maxval0:
                        maxval2 = unfiltered[j]
                    if unfiltered[j] <= minval0:
                        minval2 = unfiltered[j]
                continue
    #WRITE THE MIN AND MAX FOR EACH CENTROID   
    sheetwrite.write(lim+1,0,minval0)
    sheetwrite.write(lim+1,1,maxval0)
    sheetwrite.write(lim+2,0,minval1)
    sheetwrite.write(lim+2,1,maxval1)
    sheetwrite.write(lim+3,0,minval2)
    sheetwrite.write(lim+3,1,maxval2)
    split1 = np.array([])
    split2 = np.array([])
    split3 = np.array([])
    base1 = abs(maxval0 - minval0)/float(5)
    base2 = abs(maxval1 - minval1)/float(5)
    base3 = abs(maxval2 - minval2)/float(5)
    print(minval0)
    print(maxval0)
    #SPLIT THE RANGE FOR EACH CENTROID INTO 5 PARTS FOR RISK RATING
    for i in range(0,5):
        split1 = np.append(split1,(minval0 + ((i+1)*base1)))
        split2 = np.append(split2,(minval1 + ((i+1)*base2)))
        split3 = np.append(split3,(minval2 + ((i+1)*base3)))
    lim = 3
    print(split1)
    print(split2)
    #CHECK IF THE COMPANY FALLS IN A PARTICULAR RANGE AND PROVIDE RISK RATING
    for i in range(0,len(unfiltered)):
        if(store[i] == 0):
            if(minval0 <= unfiltered[i] <= split1[0]):
                risk = 5
            if(split1[0] <= unfiltered[i] <= split1[1]):
                risk = 4
            if(split1[1] <= unfiltered[i] <= split1[2]):
                risk = 3
            if(split1[2] <= unfiltered[i] <= split1[3]):
                risk = 2
            if(split1[3] <= unfiltered[i] <= split1[4]):
                risk = 1
        if(store[i] == 1):
            if(minval1 <= unfiltered[i] <= split2[0]):
                risk = 5
            if(split2[0] <= unfiltered[i] <= split2[1]):
                risk = 4
            if(split2[1] <= unfiltered[i] <= split2[2]):
                risk = 3
            if(split2[2] <= unfiltered[i] <= split2[3]):
                risk = 2
            if(split2[3] <= unfiltered[i] <= split2[4]):
                risk = 1
        if(store[i] == 2):
            if(minval2 <= unfiltered[i] <= split3[0]):
                risk = 5
            if(split3[0] <= unfiltered[i] <= split3[1]):
                risk = 4
            if(split3[1] <= unfiltered[i] <= split3[2]):
                risk = 3
            if(split3[2] <= unfiltered[i] <= split3[3]):
                risk = 2
            if(split3[3] <= unfiltered[i] <= split3[4]):
                risk = 1
        sheetwrite.write(lim,4,risk)
        lim = lim+1
book.save("kmeansResults.xlsx")
