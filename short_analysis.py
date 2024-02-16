import numpy as np
import pandas as pd
from bsedata.bse import BSE

b = BSE()
data = pd.read_csv("Updated_stock_data_Active.csv")

imp_columns = data[['SecurityCode', 'companyName', 'currentValue', 'change', 'pChange', 'updatedOn', 'securityID', 'group', 'faceValue', 'industry', 'previousClose', 'previousOpen', 'dayHigh', 'dayLow', '52weekHigh', '52weekLow', 'weightedAvgPrice', 'totalTradedValue', 'totalTradedQuantity', '2WeekAvgQuantity', 'marketCapFull', 'marketCapFreeFloat']]

increasing_price = []
last_five_day_up = imp_columns['previousClose'].tail(5).gt(0).all()

if last_five_day_up:
    increasing_price = imp_columns[imp_columns['previousClose'] > 0]['companyName'].tolist()
    print(f'Stock Up last 5 days {increasing_price}')
