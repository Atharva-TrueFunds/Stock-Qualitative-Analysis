import numpy as np
import pandas as pd
from bsedata.bse import BSE

data = pd.read_excel("Price_Change.xlsx", sheet_name='price_change')
bse = BSE()

symbols = data['SecurityCode'].tolist()

symbol_code = int(input("Enter the Security Code of the company you want to analyze: "))
row_values = [['1', '2', '3', '4', '5', '6', '7', '8', '9', '10']]
day_from = int(input("Enter the start date: " + str(row_values)))
print(day_from)
day_to = int(input("Enter the end date: "+ str(row_values)))
print(day_to)

if symbol_code in symbols:
    percent_change = ((day_to - day_from) / day_from) * 100
    print(f"The Percent Change for {symbol_code} from {day_from} to {day_to} is {percent_change}%")
else:
    print("Symbol not found in the dataset.")
