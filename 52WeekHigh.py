import numpy as np
import pandas as pd
from bsedata.bse import BSE
from datetime import datetime

data = pd.read_csv("Merge_currentValue.csv")

symbols = data['SecurityCode'].tolist()
b = BSE()

stock_data = []

for symbol in symbols:
    stock_quote = b.getQuote(str(symbol))
    if stock_quote:
        stock_data.append({'SecurityCode': symbol, 'currentValue': stock_quote['currentValue'],'52weekHigh': stock_quote['52weekHigh']})
    else:
        print(f"Error fetching data for symbol: {symbol}")

df = pd.DataFrame(stock_data)

df['newColumn'] = df['52weekHigh']

imp_columns = pd.concat([data['SecurityCode'], data['companyName'], df['currentValue'], df['52weekHigh'], df['newColumn']], axis=1)

if (df['currentValue'] > df['52weekHigh']).any():
    print(data.loc[df['currentValue'] > df['52weekHigh'], 'companyName'])


    # imp_columns.to_csv('Merge_currentValue.csv', index=False)
