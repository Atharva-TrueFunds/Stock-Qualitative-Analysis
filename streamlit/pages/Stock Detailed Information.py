import streamlit as st
import pandas as pd
from bsedata.bse import BSE
from datetime import datetime
from openpyxl import load_workbook
import os

def fetch_data_from_bse(scripCode):
    bse = BSE()
    stock_data = []

    for code in scripCode:
        try:
            stock_quote = bse.getQuote(str(code))
            stock_data.append(stock_quote)
        except Exception as e:
            print(f"Error fetching data for script code {code}: {e}")

    df = pd.DataFrame(stock_data)
    return df

def main():
    st.title('BSE Stock Information')

    run_button_clicked = st.button('Run')

    if run_button_clicked:
        try:
            excel_data = pd.read_excel(r"C:\Users\HP\Staging\streamlit\pages\scripCode.xlsx", engine='openpyxl')
            scripCode = excel_data['scripCode'].tolist()

            fetched_data = fetch_data_from_bse(scripCode)

            st.write(fetched_data)
            today = datetime.today().strftime('%d%b')
            sheet_name = today
            st.write(f"Saving fetched data to Excel, new sheet name: {sheet_name}...")

            with pd.ExcelWriter('stock_quote.xlsx', mode='a', engine='openpyxl') as writer:

                fetched_data.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            st.error(f"Error: {e}")
    if os.path.exists(r"C:\Users\HP\Staging\streamlit\stock_quote.xlsx"):
        if st.button("Go to stock_quote.xlsx"):
            os.startfile(r"C:\Users\HP\Staging\streamlit\stock_quote.xlsx")

        else:
            pass
if __name__ == "__main__":
    main()
