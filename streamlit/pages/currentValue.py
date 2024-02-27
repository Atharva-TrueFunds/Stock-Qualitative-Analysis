import numpy as np
import pandas as pd
from bsedata.bse import BSE
from datetime import datetime
import streamlit as st

try:
    data = pd.read_excel(r"C:\Users\HP\Staging\streamlit\stock_quote.xlsx")

    if 'scripCode' not in data.columns:
        raise KeyError("Column 'scripCode' not found in the DataFrame.")
    else:
        symbols = data['scripCode'].tolist()
        b = BSE()
        print(data.columns)

        stock_data = []

        for symbol in symbols:
            stock_quote = b.getQuote(str(symbol))
            # if stock_quote:
            stock_data.append({'scripCode': symbol,'companyName': stock_quote['companyName'],'currentValue': stock_quote['currentValue']})
            # else:
            #     print(f"Error fetching data for symbol: {symbol}")

        df = pd.DataFrame(stock_data)

        last_column_index = len(data.columns) - 1

        next_column_name = "Feb" + str(last_column_index + 26)

        data[next_column_name] = df['currentValue']

        st.write(data)

except Exception as e:
    st.error(f"Error: {e}")
