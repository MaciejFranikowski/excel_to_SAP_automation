from operator import itemgetter, attrgetter


def sortList(myList):
    for i in range(0, len(myList),1):
        temp = myList[i].split('.')
        temp[0] = int(temp[0])
        temp[1] = int(temp[1])
        temp[2] = int(temp[2])
        myList[i] = temp
    x = sorted(myList, key = itemgetter(2))
    y = sorted(x, key = lambda z: z[1])
    v = sorted(y, key = lambda z: z[0])
    for i in range(0, len(v), 1):
        temp = v[i]
        temp[0] = str(temp[0])
        temp[1] = str(temp[1])
        temp[2] = str(temp[2])
        temp = myList[i]
        temp2 = ".".join(v[i])
        v[i] = temp2
    return v
