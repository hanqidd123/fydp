# -*- coding: utf-8 -*-
import openpyxl
import os.path
import numpy as np
import sys
from sklearn import preprocessing
from sklearn.cluster import KMeans
from sklearn import datasets
import matplotlib.pyplot as plt
from sklearn.externals import joblib
from sklearn.cluster import AffinityPropagation
from sklearn import metrics
from sklearn.datasets.samples_generator import make_blobs
import random
from sklearn.cluster import MeanShift, estimate_bandwidth
from itertools import cycle

def kmeans_clustering( true_k,np_attr ):
    model = KMeans(n_clusters=true_k, init="k-means++", random_state=170, max_iter=100, n_init=1)
    model.fit(np_attr)
    clusters = model.labels_.tolist()

    order_centroids = model.cluster_centers_.argsort()[:,::-1]
    print("The cluster centroids are: \n", model.cluster_centers_)
    print("Cluster", model.labels_)

    clusters = model.labels_
    print()
    print("Sum of distances of samples to their closest cluster center: ", model.inertia_)
    return model.inertia_, model.labels_

def normalizing(np_attr):
    np_attr = np_attr.astype(np.float)
    np_attr = preprocessing.scale(np_attr)
    return np_attr

def affinity_propagation_clustering(np_attr,keys,normalized):
    # normalize data, see http://scikit-learn.org/stable/modules/preprocessing.html for differet methods
    if normalized == True:
        np_attr = normalizing(np_attr)
    #centers = [[2, 2], [8, 9], [9, 5], [3,9],[4,4],[0,0],[2,5]]
    #x, labels_true = make_blobs(n_samples=1000, centers=centers, cluster_std=0.5,random_state=0)

    #labels_true = random.sample(range(1, 1059), 1058)
    labels_true = keys
    af = AffinityPropagation(preference=-100).fit(labels_true)
    cluster_centers_indices = af.cluster_centers_indices_
    labels = af.labels_
    print(len(set(labels)))
    n_clusters_ = len(cluster_centers_indices)
    print('Estimated number of clusters: %d' % n_clusters_)
    print("Homogeneity: %0.3f" % metrics.homogeneity_score(labels_true, labels))
    print("Completeness: %0.3f" % metrics.completeness_score(labels_true, labels))
    print("V-measure: %0.3f" % metrics.v_measure_score(labels_true, labels))
    print("Adjusted Rand Index: %0.3f"
          % metrics.adjusted_rand_score(labels_true, labels))
    print("Adjusted Mutual Information: %0.3f"
          % metrics.adjusted_mutual_info_score(labels_true, labels))
    print("Silhouette Coefficient: %0.3f"
          % metrics.silhouette_score(np_attr, labels, metric='sqeuclidean'))
    # set colors for the clusters


    plt.scatter(np_attr[:, 0], np_attr[:, 1])
    plt.gray()
    plt.xlabel('X axis')
    plt.ylabel('Y axis')
    plt.show()

def meanshift_clustering(np_attr,keys,normalized):
    if normalized == True:
        np_attr = normalizing(np_attr)
    bandwidth = estimate_bandwidth(np_attr, quantile=0.2, n_samples=1058)
    ms = MeanShift(bandwidth=bandwidth, bin_seeding=True)
    ms.fit(np_attr)
    labels = ms.labels_
    cluster_centers = ms.cluster_centers_

    labels_unique = np.unique(labels)
    n_clusters_ = len(labels_unique)

    print("number of estimated clusters : %d" % n_clusters_)
    plt.figure(1)
    plt.clf()

    colors = cycle('bgrcmykbgrcmykbgrcmykbgrcmyk')
    for k, col in zip(range(n_clusters_), colors):
        my_members = labels == k
        cluster_center = cluster_centers[k]
        plt.plot(np_attr[my_members, 0], np_attr[my_members, 1], col + '.')
        plt.plot(cluster_center[0], cluster_center[1], 'o', markerfacecolor=col,
                 markeredgecolor='k', markersize=14)
    plt.title('Estimated number of clusters: %d' % n_clusters_)
    plt.show()
def output(labels,code):
    time = {}
    keys = list(code.keys())
    clusters = {}

    i = 0
    for name in code.keys():
        if labels[i] not in clusters.keys():
            clusters[labels[i]] = {}
            clusters[labels[i]]['products']= []
            clusters[labels[i]]['time'] = 0
            clusters[labels[i]]['products'].append(name)
            clusters[labels[i]]['time'] += code[name]['prod_time']

        else:
            clusters[labels[i]]['products'].append(keys[i])
            clusters[labels[i]]['time'] += code[name]['prod_time']
        i += 1

    if os.path.isfile("main.xlsx") == True:
        os.remove("main.xlsx")

    book = openpyxl.Workbook()
    sheet = book.active
    print(clusters.keys())
    for i in range(0, len(clusters)):
        sheet.cell(row=0, column=i).value = "cluster" + str(i)
        clusters[i]['time'] = clusters[i]['time']/len(clusters[i]['products'])
        for j in range(0, len(clusters[i]['products'])):
            sheet.cell(row=j + 1, column=i).value = clusters[i]['products'][j]
        sheet.cell(row=i, column=len(clusters) + 5).value = "cluster" + str(i)
        sheet.cell(row=i, column=len(clusters) + 6).value = clusters[i]['time']
    book.save("main.xlsx")





wb = openpyxl.load_workbook('data.xlsx')
data = wb.get_sheet_by_name("Sheet1")
length = 1063
i = 1

code = {}
attribute = []

while data.cell(row=i, column=0).value is not None:
    if data.cell(row=i, column=0).value not in code.keys():
        code[(data.cell(row=i, column=0).value)] = {}
        temp = []
        for k in range (2,6):
            if k == 2:
                code[(data.cell(row=i, column=0).value)]["batch_size"] = data.cell(row=i, column=k).value
                temp.append(data.cell(row=i, column=k).value)
            elif k == 3:
                code[(data.cell(row=i, column=0).value)]["prod_time"] = data.cell(row=i, column=k).value
                temp.append(data.cell(row=i, column=k).value)
            elif k == 4:
                code[(data.cell(row=i, column=0).value)]["danger_level"] = data.cell(row=i, column=k).value
                temp.append(data.cell(row=i, column=k).value)
            else:
                code[(data.cell(row=i, column=0).value)]["freq"] = data.cell(row=i, column=k).value
                temp.append(data.cell(row=i, column=k).value)

        #print(temp)
        attribute.append(temp)
    i += 1
#pprint.pprint(attribute)

keys = list(code.keys())
np_attr=np.asarray(attribute)


print(sys.argv[1])
true_k = int(sys.argv[1])
meanshift_clustering(np_attr,keys,False)
#affinity_propagation_clustering(np_attr,keys,False)
#inertia,labels = kmeans_clustering(true_k,np_attr)
#output(labels,code)

#plt.plot(initial)
#plt.show()

