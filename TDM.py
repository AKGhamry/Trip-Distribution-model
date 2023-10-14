from numpy.core.numeric import allclose
import pandas as pd
from openpyxl import load_workbook
import numpy as np
Trips1ss=pd.read_excel("InForTD.xlsx","Sheet1")
matrix=[]
print('\n')
f = numOfRows = Trips1ss.shape[0]
Trips1=Trips1ss.iloc[0:10,0:10].values.tolist()	
ProdA=Trips1ss.iloc[0:10,10:11].values
AttrA=Trips1ss.iloc[0:10,11:12].values


def sumr(Trips):
    sumR = np.zeros((f,1))
    rows = len(Trips);
    cols = len(Trips[0]);
    for i in range(0, rows):  
        sumRow= 0;  
        for j in range(0, cols):  
            sumRow = sumRow + Trips[i][j]; 
            sumR[i] = sumRow
    return sumR
def sumc(Trips11):
    sumC = np.zeros((f,1))
    rows = len(Trips11);
    cols = len(Trips11[0]);
    for i in range(0, rows):  
        sumCol = 0;  
        for j in range(0, cols):  
            sumCol = sumCol + Trips11[j][i];
            sumC[i] = sumCol
    return  sumC
def CalcFurness(Prod, Attr, Trips0):
    sumP = sum(Prod)
    sumA = sum(Attr)
    if sumP != sumA:
        Attr = Attr*(sumP/sumA)
    b = np.ones((f,1))
    a = np.ones((f,1))

    aold = bold = np.zeros((f,1))
    while allclose(b,bold,0.05) == False or allclose(a,aold,0.05) == False:
        bold = b
        aold = a
        ppp = np.transpose(Trips0)
        pfora = np.transpose(b*ppp)
        a = Prod/sumr(pfora)
        aforp = a*Trips0
        b= Attr/sumc(aforp)
        errorb = abs(((b-bold)/bold)*100)
        errora = abs(((a-aold)/aold)*100)
       
    ppp = np.transpose(Trips0)
    pfora = np.transpose(b*ppp)
    future = a*pfora
    future = np.concatenate((future,a,errora,b,errorb),axis=1)
    return future
dn=CalcFurness(ProdA, AttrA, Trips1)
df1 = pd.DataFrame((dn),
                   index=['Zone1','Zone2','Zone3','Zone4','Zone5','Zone6','Zone7','Zone8','Zone9','Zone10'],
                   columns=['Zone1','Zone2','Zone3','Zone4','Zone5','Zone6','Zone7','Zone8','Zone9','Zone10',"a's","%'ErrOf.a's","b's","%'ErrOf.b's"])
df1.to_excel("OutFromTD.xlsx",sheet_name='Sheet1',startrow=0,startcol=0)
