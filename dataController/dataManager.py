'''
Created on 7 Oct 2015

@author: Emanuele
'''

def checkData(checkList, pnFrom, pnTo, qtyFrom, qtyTo):
    '''
    check if all the element in pnFrom are in pnTo
    when an element is not found it is added in the itemList if it's not there already
    '''
    i = 0
    index = 0
    while(i in range(0,len(pnFrom))):
        if pnFrom[i] in pnTo:
            index=pnTo.index(pnFrom[i])
            if qtyFrom[i] == qtyTo[index]: 
                i=i+1
            else:
                if pnFrom[i] in checkList or qtyFrom[i] == 0:
                    i=i+1 
                else:
                    checkList.append(pnFrom[i])
                    i=i+1
        else:
            if pnFrom[i] in checkList or qtyFrom[i] == 0:
                    i=i+1 
            else:
                checkList.append(pnFrom[i])
                i=i+1
    
    return checkList

def getAllData(checkList, pnS, pnQB, descrS, descrQB, qtyS, qtyQB):
    '''
    get the description, storage quantity and quickbook quantity from the checkList
    return list of part number, description, quantity in storage, quantity in quickbook to write in the report
    in case a PN is not found, in the quantity the string "Not found" will be written
    '''
    i=0
    index = 0
    indexQB = 0
    pn=[]
    descr=[]
    qtyStorage = []
    qtyQuickbook = []
    
    while(i in range(0,len(checkList))):
        if checkList[i] in pnS and checkList[i] in pnQB:
            index = pnS.index(checkList[i])
            indexQB = pnQB.index(checkList[i])
            qtyStorage.append(qtyS[index])
            pn.append(pnQB[indexQB])
            descr.append(descrQB[indexQB])
            qtyQuickbook.append(qtyQB[indexQB])
        else:
            if checkList[i] in pnS:
                index = pnS.index(checkList[i])
                qtyStorage.append(qtyS[index])
                pn.append(pnS[index])
                descr.append(descrS[index])
                qtyQuickbook.append("Not Found")
            else:
                indexQB = pnQB.index(checkList[i])
                qtyStorage.append("Not Found")
                pn.append(pnQB[indexQB])
                descr.append(descrQB[indexQB])
                qtyQuickbook.append(qtyQB[indexQB])
            
        i=i+1
    
    return (pn,descr,qtyStorage,qtyQuickbook)