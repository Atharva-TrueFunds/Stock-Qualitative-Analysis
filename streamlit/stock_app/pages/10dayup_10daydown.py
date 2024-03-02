import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
import os

data = pd.read_excel(
    r"C:\Users\HP\Staging\streamlit\stock_app\new_sheet.xlsx", sheet_name="currentValue"
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
currentValue_data = data[
    [
        "scripCode",
        "companyName",
        "19Feb",
        "20Feb",
        "21Feb",
        "22Feb",
        "23Feb",
        "24Feb",
        "26Feb",
        "27Feb",
        "28Feb",
        "29Feb",
        "01Mar",
    ]
]

run_button_click = st.button("Run")
accending_order = []
decending_order = []

if run_button_click:
    try:
        for index, row in currentValue_data.iterrows():
            date_value = row[date.columns]

            if "19Feb" < date_value < "21Feb":
                accending_order.append(row)
    except Exception as e:
        print(f"Error: {e}")
