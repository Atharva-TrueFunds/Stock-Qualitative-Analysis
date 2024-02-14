import numpy as np
import pandas as np
import matplotlib.pyplot as plt
from bsedata.bse import BSE
# import yagmail

b= BSE()
symbol=['500570','500400','507685','500520','532439']

stock_data=[]
print("First Commit")
for symbol in symbol:
    data=b.getQuote(symbol)
    stock_data.append(data)

df=np.DataFrame(stock_data)
imp_columns= df[['companyName','currentValue','change','pChange', 'updatedOn','securityID','scripCode','group','faceValue','industry','previousClose','previousOpen','dayHigh','dayLow','52weekHigh','52weekLow','weightedAvgPrice','totalTradedValue','totalTradedQuantity','2WeekAvgQuantity','marketCapFull','marketCapFreeFloat']]
print(imp_columns)



imp_columns.to_csv('stock.csv',index= True)

# user='atharvachoudhari.truefunds@gmail.com'
# app_password='bhea sdlo jisd zoqs'
# to=''
# Subject='Stock'
# contents = ['stock.csv']
# with yagmail.SMTP(user, app_password) as yag:
#     yag.send(to, Subject, contents)
#     print('Sent email successfully')
