import numpy as np
import pandas as pd
from bsedata.bse import BSE
from datetime import datetime

data = pd.read_csv("shortdatabase.csv")

symbols = data['SecurityCode'].tolist()
b = BSE()

stock_data = []

imp_columns = pd.concat([data['SecurityCode'], data['companyName'], data['currentValue'], data['52weekHigh'], data['52weekLow']], axis=1)

data['currentValue'] = pd.to_numeric(data['currentValue'], errors='coerce')
data['currentValue'] = data['currentValue'].round(2)
data['52weekHigh'] = pd.to_numeric(data['52weekHigh'], errors='coerce')
data['52weekHigh'] = data['52weekHigh'].round(2)
companies_exceeding_52week_high = imp_columns[imp_columns['currentValue'] >= imp_columns['52weekHigh']]
companies_exceeding_52week_low = imp_columns[imp_columns['currentValue'] <= imp_columns['52weekLow']]

print(f"High{companies_exceeding_52week_high[['SecurityCode', 'companyName']]}")
print(f"Low{companies_exceeding_52week_low[['SecurityCode', 'companyName']]}")