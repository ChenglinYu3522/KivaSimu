# -*- coding: utf-8 -*-
import create
import cluster
import handle
import simu
import newHandle
import newHandle2
import handle3
import newCluster
import traceback
import randomcluster
import ani

# --------------------------------data create and ini------------------------------------------------------------
# createDate(1,orders num, itmes kinds number, ave items num, ave kind num)
ordersNumber=100
itemKindsNumber=20
aveItemsNum=5
aveKindNum=2
clusterBatchNum=6
stationNum=4
maxX=37
maxY=18
vehicleNum=20
try:
    filename=create.createDate(1,ordersNumber,itemKindsNumber,aveItemsNum,aveKindNum)


# --------------------------------no cluster method-----------------------------------------------------------
    randomcluster.createRandom(ordersNumber,stationNum,filename)
    # handle(filename,'re'+filename,orders number, items kind number, station number)
    a = handle3.handle(filename, 'random' + filename, ordersNumber, itemKindsNumber, stationNum)
    # sumu(handle[0]+2,handle[1],'final'+filename,station number,maxX,maxY,vehicleNum)
    simu.simu (a[0] + 2, a[1], 'finalrandom' + filename, stationNum, maxX, maxY, vehicleNum)
    # except:
    #     traceback.print_exc ()


# --------------------------------normal cluster method-----------------------------------------------------------
    # createSheets(filename, orders num, item kinds number, order cluster batch number , station number)
    cluster.createSheets (filename, ordersNumber, itemKindsNumber, clusterBatchNum, stationNum)
    # handle(filename,'re'+filename,orders number, items kind number, station number)
    a = handle.handle (filename, 're' + filename, ordersNumber, itemKindsNumber, stationNum)
    # sumu(handle[0]+2,handle[1],'final'+filename,station number,maxX,maxY,vehicleNum)
    simu.simu (a[0] + 2, a[1], 'final' + filename, stationNum, maxX, maxY, vehicleNum)
# except:
#     traceback.print_exc ()


# ---------------------------------new cluster method------------------------------------------------------------
    # rankArrange(filename , orders number, tiems kind number)
    newHandle.rankArrange (filename, ordersNumber, itemKindsNumber)
    # createSheets(filename, orders number ,items kind number, order cluster batch number, station number)
    newCluster.createSheets (filename, ordersNumber, itemKindsNumber, clusterBatchNum, stationNum)
    # handle(filename, 'Newre'+ filename, orders number, itemkindNumber, station number)
    b = newHandle2.handle (filename, 'Newre' + filename, ordersNumber, itemKindsNumber, stationNum)
    # sumu(handle[0]+2,handle[1],'final'+filename,station number,maxX,maxY,vehicleNum)
    simu.simu (b[0] + 2, b[1], 'newfinal' + filename, stationNum, maxX, maxY, vehicleNum)
except:
    traceback.print_exc()




