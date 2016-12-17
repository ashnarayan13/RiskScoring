import openpyxl
import numpy as np
import math
def calcStd(list1):
    sum=0.0
    avg = 0.0
    var= 0.000
    for i in range(len(list1)):
        sum=sum+list1[i]
    avg=sum/len(list1)
    print avg
    for j in range(len(list1)):
        var=var+math.pow(list1[j]-avg,2)
    var = var/len(list1)
    std = math.sqrt(var)
    print std

my_array =[1,2,3,4,5,6,7,8,9,10]
calcStd(my_array)

