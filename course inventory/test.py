'''a = [[1], [2], [3]]
b = [[1, 1], [2, 2], [3, 3]]

for i in range(len(b)):
    print(a[i] + b[i])'''

from scipy.stats import norm
rand = norm.rvs(0, 1000)
#print(rand)

import numpy
k = 1.38 * pow(10, -23)
h = 6.63 * pow(10, -34)
c = 3 * pow(10, 8)