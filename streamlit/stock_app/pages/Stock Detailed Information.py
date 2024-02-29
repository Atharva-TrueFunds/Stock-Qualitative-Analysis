import streamlit as st
import pandas as pd
from bsedata.bse import BSE
from datetime import datetime
import os

today = datetime.today().strftime("%d %b")
sheet_name = today


def fetch_data_from_bse(scripCode, sheet_name):
    bse = BSE()
    stock_data = []
    code = []
    try:
        for code in scripCode:
            try:
                stock_quote = bse.getQuote(str(code))
                fetched_data = pd.DataFrame(stock_quote)

                fetched_data["updatedOn"] = pd.to_datetime(
                    fetched_data["updatedOn"], format="%d %b %y | %I:%M %p"
                )

                fetched_data["date_part"] = fetched_data["updatedOn"].dt.strftime(
                    "%d %b"
                )

                if fetched_data["date_part"].iloc[0] != sheet_name:
                    na_data = {
                        "scripCode": [code],
                        "companyName": [fetched_data["companyName"].iloc[0]],
                        "currentValue": ["NA"],
                        "change": ["NA"],
                        "pChange": ["NA"],
                        "updatedOn": ["NA"],
                        "securityID": ["NA"],
                        "group": ["NA"],
                        "faceValue": ["NA"],
                        "industry": ["NA"],
                        "previousClose": ["NA"],
                        "previousOpen": ["NA"],
                        "dayHigh": ["NA"],
                        "dayLow": ["NA"],
                        "52weekHigh": ["NA"],
                        "52weekLow": ["NA"],
                        "weightedAvgPrice": ["NA"],
                        "totalTradedValue": ["NA"],
                        "totalTradedQuantity": ["NA"],
                        "2WeekAvgQuantity": ["NA"],
                        "marketCapFull": ["NA"],
                        "marketCapFreeFloat": ["NA"],
                        "buy": ["NA"],
                        "sell": ["NA"],
                    }
                    na_df = pd.DataFrame(na_data)
                    stock_data.append(na_df)
                else:
                    fetched_data = fetched_data.drop(columns=["date_part"])
                    stock_data.append(fetched_data.head(1))
            except Exception as e:
                st.error(f"Error fetching data for script code {code}: {e}")
        if stock_data:
            combined_data = pd.concat(stock_data, ignore_index=True)
            print(combined_data.columns)
            return combined_data
        else:
            return pd.DataFrame()

    except Exception as e:
        st.error(f"Error: {e}")


def main():
    st.title("BSE Stock Information")
    run_button_clicked = st.button("Run")

    if run_button_clicked:
        try:
            excel_data = pd.read_excel(
                r"C:\Users\HP\Staging\streamlit\stock_app\pages\scripCode.xlsx",
                engine="openpyxl",
            )
            scripCode = excel_data["scripCode"].tolist()
            fetched_data = fetch_data_from_bse(scripCode, sheet_name)
            st.write(fetched_data)
            st.write(f"Saving fetched data to Excel, new sheet name: {sheet_name}...")

            with pd.ExcelWriter(
                "stock_quote.xlsx", mode="a", engine="openpyxl"
            ) as writer:

                fetched_data.to_excel(writer, sheet_name=sheet_name, index=False)

            st.success("Data saved to stock_quote.xlsx successfully!")
        except Exception as e:
            st.error(f"Error: {e}")

        ##########    currentValue

        excel_file_path = r"C:\Users\HP\Staging\streamlit\stock_app\stock_quote.xlsx"
        current_value_file_path = (
            r"C:\Users\HP\Staging\streamlit\stock_app\currentValue.xlsx"
        )
        df_dict = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        df_current = pd.read_excel(current_value_file_path)
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
                df_stock.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            st.error(f"Error: {e}")
        from openpyxl import load_workbook

        try:
            workbook = load_workbook("currentValue.xlsx")
            source_sheet = workbook[sheet_name]
            destination_sheet = workbook["currentValue"]
            last_column_index = source_sheet.max_column
            last_column_values = []
            for row in source_sheet.iter_rows(
                min_row=1,
                max_row=source_sheet.max_row,
                min_col=last_column_index,
                max_col=last_column_index,
                values_only=True,
            ):
                for value in row:
                    last_column_values.append(value)
            next_column_index = destination_sheet.max_column + 1
            for i, value in enumerate(last_column_values, start=1):
                destination_sheet.cell(row=i, column=next_column_index, value=value)
            workbook.save("currentValue.xlsx")
        except Exception as e:
            print(f"Error: {e}")

        ##########    52weekHigh

        excel_file_path = r"C:\Users\HP\Staging\streamlit\stock_app\stock_quote.xlsx"
        weekHigh_file_path = r"C:\Users\HP\Staging\streamlit\stock_app\weekHigh.xlsx"
        df_dict = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        df_current = pd.read_excel(weekHigh_file_path)
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
                            "companyName": symbol_data["companyName"].values[0],
                            sheet_name: symbol_data["52weekHigh"].values[0],
                        }
                    )
                else:
                    print(f"No data found for symbol: {symbol}")
            df_stock = pd.DataFrame(stock_data)
            st.write(df_stock)
            st.write("Saving fetched data to Excel, new sheet name: sheet_name...")
            with pd.ExcelWriter(
                weekHigh_file_path, mode="a", engine="openpyxl"
            ) as writer:
                df_stock.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            st.error(f"Error: {e}")
        from openpyxl import load_workbook

        try:
            workbook = load_workbook("weekHigh.xlsx")
            source_sheet = workbook[sheet_name]
            destination_sheet = workbook["52weekHigh"]
            last_column_index = source_sheet.max_column
            last_column_values = []
            for row in source_sheet.iter_rows(
                min_row=1,
                max_row=source_sheet.max_row,
                min_col=last_column_index,
                max_col=last_column_index,
                values_only=True,
            ):
                for value in row:
                    last_column_values.append(value)
            next_column_index = destination_sheet.max_column + 1
            for i, value in enumerate(last_column_values, start=1):
                destination_sheet.cell(row=i, column=next_column_index, value=value)
            workbook.save("weekHigh.xlsx")
        except Exception as e:
            print(f"Error: {e}")

    if os.path.exists(r"C:\Users\HP\Staging\streamlit\stock_app\stock_quote.xlsx"):
        if st.button("Go to stock_quote.xlsx"):
            os.startfile(r"C:\Users\HP\Staging\streamlit\stock_app\stock_quote.xlsx")


if __name__ == "__main__":
    main()
