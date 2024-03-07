import pandas as pd
import streamlit as st

data = pd.read_excel(
    r"C:\Users\HP\Staging\streamlit\stock_app\code.xlsx", sheet_name="currentValue"
)
data.columns.values[2:12] = [
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
]

expected_columns = [
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
]
if not all(col in data.columns for col in expected_columns):
    st.error("Columns do not match the expected date format.")
    st.stop()

ascending_order = []
descending_order = []

# Use the correct comparison in the loop
for index, row in data.iterrows():
    dates = [row[col] for col in expected_columns]
    if dates == sorted(dates):
        ascending_order.append(row)
    elif dates == sorted(dates, reverse=True):
        descending_order.append(row)

ascending_data = pd.DataFrame(ascending_order)
descending_data = pd.DataFrame(descending_order)

st.header("Ascending Order")
st.table(ascending_data)

st.header("Descending Order")
st.table(descending_data)
