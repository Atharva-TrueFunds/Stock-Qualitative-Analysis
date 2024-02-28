import numpy as np
import pandas as pd
from bsedata.bse import BSE
from datetime import datetime
import streamlit as st

run_button_clicked = st.button("Run")
today = datetime.today().strftime("%d%b")
sheet_name = today + "_1"

excel_file_path = r"C:\Users\HP\Staging\streamlit\stock_app\stock_quote.xlsx"
current_value_file_path = r"C:\Users\HP\Staging\streamlit\stock_app\currentValue.xlsx"

df_dict = pd.read_excel(excel_file_path, sheet_name=sheet_name)
try:
    df_current = pd.read_excel(current_value_file_path)
except Exception as e:
    print(f"Error loading Excel file: {e}")
    with open("error.log", "a") as f:
        f.write(f"Error loading Excel file: {e}\n")
    # Handle the error gracefully, e.g., display an error message to the user
    st.error("Error loading Excel file. Please check the file path and sheet name.")


if run_button_clicked:
    try:
        if df_current.empty:
            symbols = df_dict["scripCode"].tolist()
            b = BSE()
            stock_data = []

            for symbol in symbols:
                symbol_data = df_dict[df_dict["scripCode"] == symbol]

                if not symbol_data.empty:
                    stock_data.append(
                        {
                            "scripCode": symbol,
                            "companyName": symbol_data["companyName"].values[0],
                            sheet_name: symbol_data["currentValue"].values[0],
                        }
                    )
                else:
                    print(f"No data found for symbol: {symbol}")

            df_stock = pd.DataFrame(stock_data)

            st.write(df_stock)
            st.write("Saving fetched data to Excel, new sheet name: sheet_name...")

            with pd.ExcelWriter(
                current_value_file_path, mode="a", engine="openpyxl"
            ) as writer:
                df_stock.to_excel(writer, sheet_name="currentValue", index=False)

        else:

            current_value_df = pd.read_excel(current_value_file_path)

            stock_quote_df = pd.read_excel(excel_file_path)

            symbols = df_dict["scripCode"].tolist()
            for symbol in symbols:
                symbol_data = df_dict[df_dict["scripCode"] == symbol]
                current_value = symbol_data["currentValue"].values[0]

                stock_quote_df[sheet_name] = current_value

            stock_quote_df.to_excel(
                current_value_file_path, index=False, sheet_name=sheet_name
            )

    except Exception as e:
        st.error(f"Error: {e}")
