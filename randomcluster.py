from openpyxl import load_workbook
import numpy as np
from openpyxl import Workbook

def createRandom(orders,Me,filename):
    wb = Workbook()
    ws = wb.active
    now=orders
    ii=0
    list=[[] for i in range(Me)]
    while now>0:
        now=now-1
        list[ii].append(now)
        ii=(ii+1)%Me

    for i in range(Me):
        ws.append(list[i])

    wb.save('random'+filename)







