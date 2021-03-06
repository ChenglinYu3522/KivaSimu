# -*- coding: utf-8 -*-
import math
from openpyxl import Workbook

#the area is a 35*39's square
def arr(maxY,maxX):
    distance=[]
    for i in range(7,maxY):
        for j in range(1,maxX):
            if j%2==0:
                distance.append([j,i,math.sqrt((j-(maxX/2))*(j-(maxX/2))+(i-7)*(i-7))])

    rank=sorted(distance, key = lambda x:x[2])

    rankPosition=[]

    for x in rank:
        rankPosition.append([x[0],x[1]])
    return rankPosition



def creatFile():
    rankPosition=arr()
    wb1 = Workbook()
    ws1 = wb1.active
    title = ['rank','x','y']
    ws1.append(title)


    for i in range(len(rankPosition)):
        content=[(i+1),rankPosition[i][0],rankPosition[i][1]]
        ws1.append(content)

    wb1.save('rankPosition.xlsx')


