# -*- coding: utf-8 -*-
import read
from openpyxl import load_workbook
import numpy as np
from openpyxl import Workbook
import traceback
#站台规划
def handle(orderFilename,reFilename,orders,items,Me):
    #读取订单
    print('reading...')
    reData=read.readResult(reFilename,Me)
    print('read finished')
    f=open('portNum'+orderFilename.strip('xlsx')+'txt','r')
    portNum=eval(f.read())
    f.close()
    f = open('rank' + orderFilename.strip('xlsx') + 'txt', 'r')
    rank = eval(f.read())
    f.close()
    f = open('ordersData' + orderFilename.strip('xlsx') + 'txt', 'r')
    ordersData = eval(f.read())
    f.close()



    ordersSq=[[] for i in range(orders)]

    for i in range(Me):
        for j in range(len(reData[i])):
            ordersSq[reData[i][j]].append(int(j/6)+1)
            ordersSq[reData[i][j]].append(i+1)


    print('rank'+orderFilename,'created')
    wb = Workbook()
    ws = wb.active
    title=['订单次序','order','station','rank','item type','item quantity','rank left']
    ws.append(title)
    #i 是orede NUMBER,j 是items num
    leftsq=[[] for i in range(items)]
    pre=[]
    pre2=[]

    for i in range(1500):
        for x in rank[i]:
            leftsq[x[0]].append(x[1])
    row=0
    maxsqe=0
    for i in range(orders):

        for j in range(items):
            if ordersData[i][j]!=0:
                mm = ordersData[i][j]
                while True:
                    content = []
                    if maxsqe<ordersSq[i][0]:
                        maxsqe=ordersSq[i][0]
                    content.append(ordersSq[i][0])
                    content.append('order' + str(i + 1))
                    content.append(ordersSq[i][1])
                    t = content
                    try:

                        if leftsq[j][0]>=mm:
                            if (len(pre2) > 2):
                                del pre2[0]
                            leftsq[j][0]=leftsq[j][0]-mm
                            t.append(portNum[j][0])
                            t.append(j)
                            t.append(mm)
                            t.append(leftsq[j][0])
                            if leftsq[j][0]==0:
                                del leftsq[j][0]
                                del portNum[j][0]
                            Fullcontent = t
                            ws.append(Fullcontent)
                            row+=1
                            break
                        else:

                            t.append(portNum[j][0])
                            t.append(j)
                            t.append(leftsq[j][0])
                            t.append(0)
                            mm=mm-leftsq[j][0]
                            Fullcontent = t
                            ws.append(Fullcontent)
                            row+=1
                            del leftsq[j][0]
                            del portNum[j][0]
                            if mm == 0:
                                break


                    except:
                        print('final'+orderFilename,'wrong!')
                        traceback.print_exc()
                        print('error')
                        return None
    wb.save('newfinal'+orderFilename)
    return (row,maxsqe)
    print('Newfinal'+orderFilename,'created')
    print('---------------------------------one group end--------------------------------------')
    print()
    print()
    print()




