import numpy as np
import pandas as pd
from bsedata.bse import BSE
from datetime import datetime

data = pd.read_csv("Merge_currentValue.csv")  # Load existing data

symbols = data['SecurityCode'].tolist()
b = BSE()

stock_data = []

for symbol in symbols:
    stock_quote = b.getQuote(str(symbol))
    if stock_quote:
        stock_data.append({'SecurityCode': symbol, 'currentValue': stock_quote['currentValue']})
    else:
        print(f"Error fetching data for symbol: {symbol}")

df = pd.DataFrame(stock_data)

last_column_index = len(data.columns) - 1

next_column_name = chr(ord('A') + last_column_index)

data[next_column_name] = df['currentValue']

data.to_csv('Merge_currentValue.csv', index=False)
