import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from bsedata.bse import BSE

b = BSE()
data = pd.read_excel("imp_stock_quote_copy.xlsx", sheet_name="23Feb")

symbol = data["scripCode"].tolist()
print(symbol)

stock_data = []

for symbol in symbol:
    try:
        stock_quote = b.getQuote(str(symbol))
        stock_data.append(stock_quote)
    except Exception as e:
        pass

df = pd.DataFrame(stock_data)
imp_columns = df[
    [
        "scripCode",
        "companyName",
        "currentValue",
        "change",
        "pChange",
        "updatedOn",
        "securityID",
        "group",
        "faceValue",
        "industry",
        "previousClose",
        "previousOpen",
        "dayHigh",
        "dayLow",
        "52weekHigh",
        "52weekLow",
        "weightedAvgPrice",
        "totalTradedValue",
        "totalTradedQuantity",
        "2WeekAvgQuantity",
        "marketCapFull",
        "marketCapFreeFloat",
    ]
]
print(imp_columns)

imp_columns.to_excel("imp_stock_quote_copy.xlsx", sheet_name="23Feb", index=True)
