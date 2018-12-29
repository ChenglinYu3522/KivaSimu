# -*- coding: utf-8 -*-
import random

from openpyxl import Workbook

#input the num to create data
# sheets=int(input('num of the sheets'))
# orders=int(input('num of orders for each sheets'))
# items=int(input('the kinks of items'))
# aveitems=float(input('aveitems for each order'))
# avekinds=float(input('avekinds for each order'))
# headname=str(input('input the headfilename of files'))
# filenum=1

#数据模拟
def createDate(sheets,orders,items,aveitems,avekinds):
    print('orders,items,aveitems,avekinds:',orders,items,aveitems,avekinds,'data begin create')
    filenum=1;
    while(True):
        #open a memory to store the data
        wb = Workbook()
        ws = wb.active
        print('create new xls scuss')
        #caculate the sumitems and sumkinds
        sumitems=int(orders*aveitems)
        sumkinds=int(orders*avekinds)

        #to creat the sumkinds of num
        spit=[False for i in range(sumitems)]
        spit[0]=True
        j=1;
        i=0;
        x=[]
        x.append(0)
        while(j<int(sumkinds)):
            i=(i+random.randint(1,sumitems))%sumitems
            if(not spit[i]):
                spit[i]=True
                x.append(i)
                j=j+1;
        x=sorted(x)
        rd=[]
        for i1 in range(len(x)-1):
            k=x[i1+1]-x[i1]
            rd.append(k)
        rd.append(sumitems-x[sumkinds-1])
        #to choose wehre in sheets to place the data
        biaoge=[False for i2 in range(items*orders+1)]

        a=0
        sss={}
        ord=orders
        for i3 in range(orders):
            t=random.randint(a+1,a+items)
            biaoge[t]=True;
            sss[str(t)]=rd[i3]
            a=a+items
        a=0
        try:
            for i3 in range(items):
                flag=False
                for i4 in range(orders):
                    if(biaoge[i3+1+items*i4]):
                        flag=True;
                        break;
                if(not flag):
                    t=random.randint(0,orders-1)
                    z=i3+1+t*items
                    biaoge[z]=True
                    sss[str(z)]=rd[ord]
                    ord+=1
                    if (ord >= sumitems):
                        break
        except:
            print('---****--------****-----')

        print('create data scuss')
        for i4 in range(ord,sumkinds):
            t = random.randint(1, items*orders)
            while(biaoge[t]):
                t=random.randint(1,items*orders)
            biaoge[t]=True
            sss[str(t)] = rd[i4]

        print('begin to put data into new xls...')
        #input the data to sheets
        pp1=['']
        for j in range(1,orders+1):
            pp1.append('order'+str(j))
        ws.append(pp1)

        for i5 in range(1,items+1):
            pp=['item'+str(i5)]
            for j1 in range(1, orders+1):
                z=items*(j1-1)+i5
                if str(z) in sss:
                    pp.append(int(sss[str(z)]))
                else:
                     pp.append(0)
            ws.append(pp)
        print('save scuss')
        #save the sheets
        filename1='o'+str(orders)+'tc'+str(items)+'ac'+str(avekinds)+'an'+str(aveitems)+'.xlsx'
        wb.save(filename1)
        print(filename1,'created!')
        filenum+=1
        if(filenum>sheets):
            return filename1
            break;



