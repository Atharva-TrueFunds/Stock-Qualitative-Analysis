import pandas as pd
import streamlit as st

data = pd.read_excel(
    r"C:\Users\HP\Staging\streamlit\stock_app\All_sheets.xlsx",
    sheet_name="currentValue",
)
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
date_columns = [
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

if not set(date_columns).issubset(data.columns):
    st.error("Columns do not match the expected date format.")
    st.stop()

ascending_mask = data[date_columns].apply(
    lambda row: all(row[i] <= row[i + 1] for i in range(len(row) - 1)), axis=1
)
descending_mask = data[date_columns].apply(
    lambda row: all(row[i] >= row[i + 1] for i in range(len(row) - 1)), axis=1
)

ascending_data = data[ascending_mask]
descending_data = data[descending_mask]

st.header("Ascending Order")
st.table(ascending_data)

st.header("Descending Order")
st.table(descending_data)
