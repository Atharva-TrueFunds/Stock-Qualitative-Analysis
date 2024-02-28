import numpy as np
import pandas as pd
from bsedata.bse import BSE
from datetime import datetime
import streamlit as st
import os

run_button_clicked = st.button("Run")
today = datetime.today().strftime("%d%b")
sheet_name = today

excel_file_path = r"C:\Users\HP\Staging\streamlit\stock_app\stock_quote.xlsx"
existing_excel_path = r"C:\Users\HP\Staging\streamlit\stock_app\currentValue.xlsx"

df_dict = pd.read_excel(
    excel_file_path, sheet_name="currentValue"
)  # Read from the existing sheet "28Feb"
df_empty = pd.read_excel(existing_excel_path)

if run_button_clicked:
    if df_empty.empty:
        try:
            symbols = df_dict["scripCode"].tolist()
            b = BSE()
            stock_data = []

            for symbol in symbols:
                symbol_data = df_dict[df_dict["scripCode"] == symbol]

                if not symbol_data.empty:
                    stock_data.append(
                        {
                            "scripCode": symbol,
                            sheet_name: symbol_data["currentValue"].values[0],
                        }
                    )
                else:
                    print(f"No data found for symbol: {symbol}")

            df_stock = pd.DataFrame(stock_data)

            # Write the new data as a new column in the existing Excel sheet
            with pd.ExcelWriter(
                existing_excel_path, mode="a", engine="openpyxl"
            ) as writer:
                df_stock.to_excel(writer, sheet_name="currentValue", index=False)

            st.write("Fetched data appended to Excel file.")

        except Exception as e:
            st.error(f"Error: {e}")
