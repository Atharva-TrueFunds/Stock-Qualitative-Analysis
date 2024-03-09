import numpy as np
import pandas as pd
import streamlit as st

data = pd.read_excel("All_sheets.xlsx", sheet_name="currentValue")

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

date_columns = data.columns[2:16]
currentValue_data = data[["scripCode"] + ["companyName"] + list(date_columns)]

highest_values = []
highest_column_names = []
lowest_values = []
lowest_column_names = []

for index, row in currentValue_data.iterrows():
    numeric_values = pd.to_numeric(row[date_columns], errors="coerce")
    highest_values.append(numeric_values.max())
    highest_column_index = np.argmax(numeric_values)
    highest_column_names.append(date_columns[highest_column_index])
    lowest_values.append(numeric_values.min())
    lowest_column_index = np.argmin(numeric_values)
    lowest_column_names.append(date_columns[lowest_column_index])


currentValue_data["Highest Value"] = highest_values
currentValue_data["Highest Value Column"] = highest_column_names
currentValue_data["Lowest Value"] = lowest_values
currentValue_data["Lowest Value Column"] = lowest_column_names

st.table(currentValue_data)
