import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from bsedata.bse import BSE

b= BSE()
data=pd.read_csv("Equity.csv")

symbol= data['Security Code'].tolist()
print(symbol)

stock_data=[]

for symbol in symbol:
    try:
        stock_quote=b.getQuote(str(symbol))
        stock_data.append(stock_quote)
    except Exception as e:
        pass

df=pd.DataFrame(stock_data)
imp_columns= df[['Security Code','companyName','currentValue','change','pChange', 'updatedOn','securityID','scripCode','group','faceValue','industry','previousClose','previousOpen','dayHigh','dayLow','52weekHigh','52weekLow','weightedAvgPrice','totalTradedValue','totalTradedQuantity','2WeekAvgQuantity','marketCapFull','marketCapFreeFloat']]
print(imp_columns)

# imp_columns.to_csv('Updated_stock_data_Active.csv',index= True)