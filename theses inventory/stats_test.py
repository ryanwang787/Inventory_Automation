import numpy as np
import matplotlib.pyplot as plt

x = np.arange(0, 24, 1)
y = np.array([3.3, 3.6, 4.7, 6.3, 8.1, 9.6, 10.5, 10.6, 9.8, 8.2, 6.3, 4.6, 3.4, 3.1, 3.7, 5.1, 6.8, 8.4, 9.7, 10.2, 9.9, 8.8, 7.1, 5.4])

a = -3.75
b = 0.5
h = 0
k = 6.75

cos = a * np.cos(b * (x - h)) + k

fig = plt.figure(figsize=(12,8))
plt.plot(x, y, 'ok')
plt.plot(x, cos, 'r', linestyle='-', marker='o')
plt.show()


