import csv
import numpy as np
import pylab as pl
import scipy.stats as stats


def calcVar(filename):
    daxFile = open(filename)
    reader = csv.reader(daxFile)
    data = list(reader)
    current = []
    i = 0
    for d in data:
        if (i < 2):
            i = i + 1
            continue
        l = d[0].__len__()
        if (l == 0):
            continue
        current.append(float(d[0][0:l - 1]))
        i = i + 1
    fit = stats.norm.pdf(current, np.mean(current), np.std(current))
    pl.plot(current, fit, '-o')
    pl.hist(current, normed=True)
    print np.percentile(current, q=5)
    pl.show()

fileName=raw_input("Enter file name :")
calcVar(fileName)

