# -*- coding: utf-8 -*-
import read
from openpyxl import load_workbook
import numpy as np
from openpyxl import Workbook
import traceback
def rankArrange(orderFilename,orders,items):
    #读取订单
    print('reading...')
    ordersData=read.readOrders(orderFilename,orders,items)
    #reData=read.readResult(reFilename)
    print('read finished')

    #统计item出现频次
    itemsAp=[[0,i] for i in range(items)]
    room=[0 for i in range(items)]
    for i in range(items):
        for j in range(orders):
            if ordersData[j][i]!=0:
                itemsAp[i][0]+=1
                room[i]+=ordersData[j][i]




    portNum=[[] for i in range(items)]
    #对出现频次进行降序
    newI=sorted(itemsAp, key=lambda itemsAp: itemsAp[0],reverse=True)
    rank=[[] for i in range(1500)]

    #进行架号编写前10%，10%-30%，余下
    ordersSq=[[] for i in range(orders)]
    portN=1
    s=0
    left=400
    same=[]
    for i in range(0,int(items*0.2),2):
        sumit=room[newI[i][1]]
        sumit2=room[newI[i+1][1]]
        same.append((newI[i][1],newI[i+1][1]))
        if (sumit % 100 == 0):
            sumit = int(sumit / 100) * 100
        else:
            sumit = (int(sumit / 100) + 1) * 100
        if (sumit2 % 100 == 0):
            sumit2 = int(sumit2 / 100) * 100
        else:
            sumit2 = (int(sumit2 / 100) + 1) * 100
        if (sumit+sumit2)>=left:
            while (sumit+sumit2)>=left:
                portNum[newI[i][1]].append(portN)
                portNum[newI[i+1][1]].append(portN)
                rank[portN].append((newI[i][1],left/2))
                rank[portN].append((newI[i+1][1], left / 2))
                portN+=1
                sumit=sumit-left/2
                sumit2=sumit2-left/2
                left=400

            if (sumit+sumit2)==0:
                continue
            else:
                maxm=sumit2
                if sumit>sumit2:
                    maxm=sumit

                left=left-maxm*2
                portNum[newI[i][1]].append(portN)
                portNum[newI[i+1][1]].append(portN)
                rank[portN].append((newI[i][1],maxm))
                rank[portN].append((newI[i+1][1], maxm))

        else:
            portNum[newI[i][1]].append(portN)
            portNum[newI[i + 1][1]].append(portN)
            rank[portN].append((newI[i][1], sumit))
            rank[portN].append((newI[i + 1][1], sumit2))
            left = left - (sumit + sumit2)

    for i in range(int(items*0.2),items):

        sumit=room[newI[i][1]]
        if (sumit % 100 == 0):
            sumit = int(sumit / 100) * 100
        else:
            sumit = (int(sumit / 100) + 1) * 100
        if sumit>=left:
            while sumit>=left:
                portNum[newI[i][1]].append(portN)
                rank[portN].append((newI[i][1],left))
                portN+=1
                sumit=sumit-left
                left=400

            if sumit==0:
                continue
            else:
                left=left-sumit
                portNum[newI[i][1]].append(portN)
                rank[portN].append((newI[i][1],sumit))
        else:
            portNum[newI[i][1]].append(portN)
            rank[portN].append((newI[i][1], sumit))
            left = left - sumit
    print('arrange rank finished')
    wb1 = Workbook()
    ws1 = wb1.active
    title = ['rank','port','type','number']
    ws1.append(title)

    for i in range(1500):
        for x in rank[i]:
            content = []
            content.append(i)
            content.append(i)
            content.append(x[0])
            content.append(x[1])
            ws1.append(content)
    f=open('portNum'+orderFilename.strip('xlsx')+'txt','w')
    f.write(str(portNum))
    f.close()
    f = open('rank' + orderFilename.strip('xlsx') + 'txt', 'w')
    f.write(str(rank))
    f.close()
    f = open('same' + orderFilename.strip('xlsx') + 'txt', 'w')
    f.write(str(same))
    f.close()
    f = open('ordersData' + orderFilename.strip('xlsx') + 'txt', 'w')
    f.write(str(ordersData))
    f.close()
    wb1.save('Newrank'+orderFilename)