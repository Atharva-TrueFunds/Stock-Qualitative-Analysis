import numpy as np
import pandas as pd
from bsedata.bse import BSE
from datetime import datetime

data = pd.read_csv("shortdatabase.csv")

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

# Filter companies where current value exceeds 52-week high
companies_exceeding_52week_high = imp_columns[imp_columns['currentValue'] > imp_columns['52weekHigh']]

print(companies_exceeding_52week_high[['SecurityCode', 'companyName']])
