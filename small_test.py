import numpy
list = numpy.zeros((5,3))
list[1][1] = 1


elist = [2,2,2]

for i in range(3):
    list[1][i] = elist[i]

print(list)