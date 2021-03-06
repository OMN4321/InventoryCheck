'''
Created on 2 Oct 2015

@author: Emanuele
'''

from fileController import fileManager
from dataController import dataManager  
import sys


'''for LabView'''
storagePath = str(sys.argv[1])
quickbookPath =str(sys.argv[2])
reportPath =str(sys.argv[3])

'''Python debug'''
# storagePath = "X://UtilityFiles//Storage.xlsx"
# quickbookPath = "C://Users//emanu//Desktop//COSMED Asia-Pacific Pty Ltd_ProductServiceList.xlsx"
# reportPath = "C://Users//emanu//Desktop//report.xlsx" 


pnQuickbook, descrQuickbook, qtyQuickbook = fileManager.getData(quickbookPath, "notstorage")

pnStorage, descrStorage, qtyStorage = fileManager.getData(storagePath, "storage")

itemList = []
checkList = []
pn = []
description = []
qtyS = []
qtyQB = []

checkItems = dataManager.checkData(itemList, pnStorage, pnQuickbook, qtyStorage, qtyQuickbook)
checkItems = dataManager.checkData(checkItems, pnQuickbook, pnStorage, qtyQuickbook, qtyStorage)

pn, description, qtyS, qtyQB = dataManager.getAllData(checkItems, pnStorage, pnQuickbook, descrStorage, descrQuickbook, qtyStorage, qtyQuickbook)


report = fileManager.createFile(reportPath)
fileManager.writeInto(reportPath, pn, description, qtyS, qtyQB)
missmatch = fileManager.columnAdjustment(reportPath)
if missmatch:
    print('You need to double check some items in the inventory in the file:', reportPath)
else:
    print('all good!!!')
