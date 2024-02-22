# import numpy as np
# import pandas as pd
# from bsedata.bse import BSE
# from datetime import datetime

# data = pd.read_csv("MarketCap_4pm.csv")

# symbols = data['SecurityCode'].tolist()
# b = BSE()

# stock_data = []

# for symbol in symbols:
#     stock_quote = b.getQuote(str(symbol))
#     if stock_quote:
#         stock_data.append({'SecurityCode': symbol, 'marketCapFull': stock_quote['marketCapFull']})
#     else:
#         print(f"Error fetching data for symbol: {symbol}")

# df = pd.DataFrame(stock_data)

# last_column_index = len(data.columns) - 1

# next_column_name = "Feb" + str(last_column_index + 16)

# data[next_column_name] = df['marketCapFull']

# data.to_csv('MarketCap_4pm.csv', index=False)



from nsetools import Nse
from pprint import pprint

nse = Nse()

quote = nse.get_quote("INFY")
pprint(quote)
