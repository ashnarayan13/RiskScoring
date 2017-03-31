from pylab import plot,show
from numpy import vstack,array
from numpy.random import rand
from scipy.cluster.vq import kmeans,vq

# data generation
data = vstack((rand(400,2) + array([.5,.5]),rand(400,2)))
print(data)
# computing K-Means with K = 2 (2 clusters)
centroids,_ = kmeans(data,3)
# assign each sample to a cluster
idx,_ = vq(data,centroids)

# some plotting using numpy's logical indexing
plot(data[idx==0,0],data[idx==0,1],'ob',
     data[idx==1,0],data[idx==1,1],'or',
     data[idx==2,0],data[idx==2,1],'og') # third cluster points
plot(centroids[:,0],centroids[:,1],'sm',markersize=8)
show()
