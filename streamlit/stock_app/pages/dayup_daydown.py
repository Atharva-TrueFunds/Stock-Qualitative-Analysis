import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
import os

data = pd.read_excel(
    r"C:\Users\HP\Staging\streamlit\stock_app\All_sheets.xlsx", sheet_name="master"
)
print(data.columns)

data.columns.values[2:12] = [
    "2024-02-19 00:00:00",
    "2024-02-20 00:00:00",
    "2024-02-21 00:00:00",
    "2024-02-22 00:00:00",
    "2024-02-23 00:00:00",
    "2024-02-26 00:00:00",
    "2024-02-27 00:00:00",
    "2024-02-28 00:00:00",
    "2024-02-29 00:00:00",
    "2024-03-01 00:00:0",
]

currentValue_data = data[
    [
        "scripCode",
        "2024-02-19 00:00:00",
        "2024-02-20 00:00:00",
        "2024-02-21 00:00:00",
        "2024-02-22 00:00:00",
        "2024-02-23 00:00:00",
        "2024-02-26 00:00:00",
        "2024-02-27 00:00:00",
        "2024-02-28 00:00:00",
        "2024-02-29 00:00:00",
        "2024-03-01 00:00:0",
    ]
]

run_button_click = st.button("Run")
ascending_order = []
descending_order = []

if run_button_click:
    try:
        for index, row in currentValue_data.iterrows():
            for date_column in currentValue_data.columns[1:]:
                date_value = row[date_column]

                if "19Feb" < date_value < "01Mar":
                    ascending_order.append(row)
                    st.table(pd.DataFrame(ascending_order))
    except Exception as e:
        print(f"Error: {e}")

# if __name__ == "__main__":
#     st.run()
