# coding=utf-8
import numpy as np
import random
from sklearn.cluster import KMeans


def create_data(sheet, col):
    '''Read and store the data as two lists. One which is used to create duplicates and one
    for computation
    sheet = xlrd worksheet to extract the data
    col = column in the sheet where the data is'''
    original_vals = np.array([])
    km_vals = np.array([])
    for i in range(2, sheet.nrows):
        temp = float(sheet.cell_value(i, col))
        original_vals = np.append(original_vals, temp)
        km_vals = np.append(km_vals, temp)
    return original_vals, km_vals


def kmeans_setup(data):
    '''Use numpy random to add random values uniformly across the data
    then compute kmeans clusters
    data = input data as numpy list'''
    avg = (max(data) - min(data)) / 16
    print(min(data), max(data))
    mins = min(data) - avg
    maxs = max(data) + avg
    for vals in range(10000):
        temp = random.uniform(mins, maxs)
        data = np.append(data, temp)
    data = data.reshape(data.size, -1)
    kmeans = KMeans(n_clusters=3, random_state=0, max_iter=300, n_init=100, algorithm='auto').fit(data)
    # temp = kmeans.predict(data)
    # scipy.stats.mstats.normaltest(temp, axis=0)
    return kmeans, data


def map_distance(kmeans, data, vals):
    '''Get cluster centers from kmeans and get the split values to calculate the risk
    kmeans = kmeans calculated from the kmeans_setup()
    data = the annotated values with random numbers
    vals = original data'''
    bins = kmeans.predict(data)
    distances = np.array([])
    mappings = np.array([])
    min_max = np.array([])
    print("cluster centers")
    actual = np.array([])
    actual = np.append(actual, min(kmeans.cluster_centers_))
    for i in kmeans.cluster_centers_:
        if i != max(kmeans.cluster_centers_) and i != min(kmeans.cluster_centers_):
            actual = np.append(actual, i)
    actual = np.append(actual, max(kmeans.cluster_centers_))
    maxval0 = float(actual[0])
    minval0 = float(actual[0])
    maxval1 = float(actual[1])
    minval1 = float(actual[1])
    maxval2 = float(actual[2])
    minval2 = float(actual[2])
    nulls = 0
    ones = 0
    twos = 0
    for i in range(0, len(vals)):
        count = 0
        for j in range(0, len(data)):
            if vals[i] == data[j] and count == 0:
                count = 1
                if bins[i] == 0:
                    nulls += 1
                elif bins[i] == 1:
                    ones += 1
                elif bins[i] == 2:
                    twos += 1
                mappings = np.append(mappings, kmeans.cluster_centers_[bins[i]])
                temp_dist = vals[i] - kmeans.cluster_centers_[bins[i]]
                distances = np.append(distances, temp_dist)
                # associate min max for each cĺuster
                if kmeans.cluster_centers_[bins[i]] == actual[0]:
                    if vals[i] >= maxval0:
                        maxval0 = vals[i]
                    if vals[i] <= minval0:
                        minval0 = vals[i]
                elif kmeans.cluster_centers_[bins[i]] == actual[1]:
                    if vals[i] >= maxval1:
                        maxval1 = vals[i]
                    if vals[i] <= minval1:
                        minval1 = vals[i]
                elif kmeans.cluster_centers_[bins[i]] == actual[2]:
                    if vals[i] >= maxval2:
                        maxval2 = vals[i]
                    if vals[i] <= minval2:
                        minval2 = vals[i]
    print(nulls)
    print(ones)
    print(twos)
    min_max = np.append(min_max, minval0)
    min_max = np.append(min_max, maxval0)
    min_max = np.append(min_max, minval1)
    min_max = np.append(min_max, maxval1)
    min_max = np.append(min_max, minval2)
    min_max = np.append(min_max, maxval2)
    return mappings, distances, min_max, actual


def risk_vals(ranges_final, vals, sorted_pts, mapping, type):
    '''Calculate the risk values of the data
    ranges_final = the min and max value of the clusters
    vals = original data
    sorted_pts = The centers of the clusters
    mapping = mapping of data to corresponding cluster
    type = the scoring pattern'''
    split1 = np.array([])
    split2 = np.array([])
    split3 = np.array([])
    base1 = float(ranges_final[1] - ranges_final[0]) / float(5)
    base2 = float(ranges_final[3] - ranges_final[2]) / float(5)
    base3 = float(ranges_final[5] - ranges_final[4]) / float(5)
    split1 = np.append(split1, ranges_final[0])
    split2 = np.append(split2, ranges_final[2])
    split3 = np.append(split3, ranges_final[4])
    for pts in range(0, 4):
        split1 = np.append(split1, (ranges_final[0] + ((pts + 1) * base1)))
        split2 = np.append(split2, (ranges_final[2] + ((pts + 1) * base2)))
        split3 = np.append(split3, (ranges_final[4] + ((pts + 1) * base3)))
    split1 = np.append(split1, ranges_final[1])
    split2 = np.append(split2, ranges_final[3])
    split3 = np.append(split3, ranges_final[5])
    print("split1")
    for j in split1:
        print(int(j))
    print("split2")
    for j in split2:
        print(int(j))
    print("split3")
    for j in split3:
        print(int(j))
    # Change the risk numbers based on the parameter
    risk = np.array([])
    if type==1:
        for i in range(0, len(vals)):
            if mapping[i] == sorted_pts[0]:
                if split1[0] <= vals[i] <= split1[1]:
                    risk = np.append(risk, 1)
                elif split1[1] <= vals[i] <= split1[2]:
                    risk = np.append(risk, 2)
                elif split1[2] <= vals[i] <= split1[3]:
                    risk = np.append(risk, 3)
                elif split1[3] <= vals[i] <= split1[4]:
                    risk = np.append(risk, 4)
                elif split1[4] <= vals[i] <= split1[5]:
                    risk = np.append(risk, 5)
            elif mapping[i] == sorted_pts[1]:
                if split2[0] <= vals[i] <= split2[1]:
                    risk = np.append(risk, 1)
                elif split2[1] <= vals[i] <= split2[2]:
                    risk = np.append(risk, 2)
                elif split2[2] <= vals[i] <= split2[3]:
                    risk = np.append(risk, 3)
                elif split2[3] <= vals[i] <= split2[4]:
                    risk = np.append(risk, 4)
                elif split2[4] <= vals[i] <= split2[5]:
                    risk = np.append(risk, 5)
            elif mapping[i] == sorted_pts[2]:
                if split3[0] <= vals[i] <= split3[1]:
                    risk = np.append(risk, 1)
                elif split3[1] <= vals[i] <= split3[2]:
                    risk = np.append(risk, 2)
                elif split3[2] <= vals[i] <= split3[3]:
                    risk = np.append(risk, 3)
                elif split3[3] <= vals[i] <= split3[4]:
                    risk = np.append(risk, 4)
                elif split3[4] <= vals[i] <= split3[5]:
                    risk = np.append(risk, 5)
    elif type==2:
        for i in range(0, len(vals)):
            if mapping[i] == sorted_pts[0]:
                if split1[0] <= vals[i] <= split1[1]:
                    risk = np.append(risk, 5)
                elif split1[1] <= vals[i] <= split1[2]:
                    risk = np.append(risk, 4)
                elif split1[2] <= vals[i] <= split1[3]:
                    risk = np.append(risk, 3)
                elif split1[3] <= vals[i] <= split1[4]:
                    risk = np.append(risk, 2)
                elif split1[4] <= vals[i] <= split1[5]:
                    risk = np.append(risk, 1)
            elif mapping[i] == sorted_pts[1]:
                if split2[0] <= vals[i] <= split2[1]:
                    risk = np.append(risk, 5)
                elif split2[1] <= vals[i] <= split2[2]:
                    risk = np.append(risk, 4)
                elif split2[2] <= vals[i] <= split2[3]:
                    risk = np.append(risk, 3)
                elif split2[3] <= vals[i] <= split2[4]:
                    risk = np.append(risk, 2)
                elif split2[4] <= vals[i] <= split2[5]:
                    risk = np.append(risk, 1)
            elif mapping[i] == sorted_pts[2]:
                if split3[0] <= vals[i] <= split3[1]:
                    risk = np.append(risk, 5)
                elif split3[1] <= vals[i] <= split3[2]:
                    risk = np.append(risk, 4)
                elif split3[2] <= vals[i] <= split3[3]:
                    risk = np.append(risk, 3)
                elif split3[3] <= vals[i] <= split3[4]:
                    risk = np.append(risk, 2)
                elif split3[4] <= vals[i] <= split3[5]:
                    risk = np.append(risk, 1)
    elif type==3:
        for i in vals:
            if i <= 0.25 and i >= 0:
                risk = np.append(risk, 1)
            elif i <= 0.5 and i >= 0.25:
                risk = np.append(risk, 2)
            elif i <= 0.75 and i >= 0.5:
                risk = np.append(risk, 3)
            elif i <= 1 and i >= 0.75:
                risk = np.append(risk, 4)
            else:
                risk = np.append(risk, 5)
    print(len(vals))
    print(len(risk))
    return risk


def detect_duplicates(sheet_read):
    '''Detect duplicate values in the dataset as we fill missing values
    sheet_read = sheet containing the data'''
    temp_ctr = 0
    for el in set(sheet_read):
        if sheet_read.tolist().count(el) > 1:
            temp_ctr += sheet_read.tolist().count(el)
            # print el, sheet_read.count(el), len(sheet_read) - x[::-1].index(el)
    print(temp_ctr)
    return temp_ctr
# END OF FUNCTIONS

