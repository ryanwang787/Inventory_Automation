import matplotlib.pyplot as plt
import numpy as np

years = ['2019-2020', '2020-2021', '2021-2022']
sust_courses = [2228, 2233, 3059]
sust_courses_txt = ['2228', '2233', '3059']
total_courses = [7840, 8521, 10366]

fig = plt.figure(figsize=(12, 8))
plt.plot(years, total_courses)
plt.scatter(years, total_courses)
for i in range(3):
    plt.annotate(sust_courses_txt[i], (years[i], total_courses))
plt.bar(years, height=sust_courses, width=0.25, color='b')
plt.yticks(np.arange(0, 11000, 1000))
plt.show()