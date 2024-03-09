import pandas as pd
import streamlit as st

data = pd.read_excel("All_sheets.xlsx", sheet_name="52weekHigh")
data.columns.values[2:16] = [
    "19Feb",
    "20Feb",
    "21Feb",
    "22Feb",
    "23Feb",
    "26Feb",
    "27Feb",
    "28Feb",
    "29Feb",
    "01Mar",
    "02Mar",
    "05Mar",
    "06Mar",
    "07Mar",
]

result_rows = []

for index, row in data.iterrows():
    scrip_code = row["scripCode"]
    company_name = row["companyName"]

    last_day_value = row.iloc[-1]

    for i in range(2, 16):
        if row.iloc[i] == last_day_value:
            result_rows.append(
                {
                    "scripCode": scrip_code,
                    "companyName": company_name,
                    "day": data.columns.values[i],
                }
            )
            break

result_df = pd.DataFrame(result_rows)

st.table(result_df)
