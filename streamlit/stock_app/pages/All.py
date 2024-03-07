import streamlit as st
import pandas as pd

st.sidebar.title("All_sheets")
page = st.sidebar.radio(
    "Go to",
    [
        "currentValue",
        "Change",
        "pChange",
        "dayHigh",
        "dayLow",
        "previousClose",
        "previousOpen",
        "52weekHigh",
        "52weekLow",
        "WeightedAvgPrice",
        "2WeekAvgQuantity",
        "totalTradedQuantity",
        "totalTradedValue",
        "marketCapFreeFloat",
        "marketCapFull",
    ],
)

excel_file = "All_sheets.xlsx"

if page == "currentValue":
    st.title("currentValue")
    df_currentValue = pd.read_excel(excel_file, sheet_name="currentValue")
    st.write(df_currentValue)

elif page == "Change":
    st.title("Change")
    df_Change = pd.read_excel(excel_file, sheet_name="Change")
    st.write(df_Change)

elif page == "pChange":
    st.title("pChange")
    df_pChange = pd.read_excel(excel_file, sheet_name="pChange")
    st.write(df_pChange)

elif page == "dayHigh":
    st.title("dayHigh")
    df_dayHigh = pd.read_excel(excel_file, sheet_name="dayHigh")
    st.write(df_dayHigh)

elif page == "dayLow":
    st.title("dayLow")
    df_dayLow = pd.read_excel(excel_file, sheet_name="dayLow")
    st.write(df_dayLow)

elif page == "previousClose":
    st.title("previousClose")
    df_previousClose = pd.read_excel(excel_file, sheet_name="previousClose")
    st.write(df_previousClose)

elif page == "previousOpen":
    st.title("previousOpen")
    df_previousOpen = pd.read_excel(excel_file, sheet_name="previousOpen")
    st.write(df_previousOpen)

elif page == "52weekHigh":
    st.title("52weekHigh")
    df_52weekHigh = pd.read_excel(excel_file, sheet_name="52weekHigh")
    st.write(df_52weekHigh)

elif page == "52weekLow":
    st.title("52weekLow")
    df_52weekLow = pd.read_excel(excel_file, sheet_name="52weekLow")
    st.write(df_52weekLow)

elif page == "WeightedAvgPrice":
    st.title("WeightedAvgPrice")
    df_WeightedAvgPrice = pd.read_excel(excel_file, sheet_name="WeightedAvgPrice")
    st.write(df_WeightedAvgPrice)

elif page == "2WeekAvgQuantity":
    st.title("2WeekAvgQuantity")
    df_2WeekAvgQuantity = pd.read_excel(excel_file, sheet_name="2WeekAvgQuantity")
    st.write(df_2WeekAvgQuantity)

elif page == "totalTradedQuantity":
    st.title("totalTradedQuantity")
    df_totalTradedQuantity = pd.read_excel(excel_file, sheet_name="totalTradedQuantity")
    st.write(df_totalTradedQuantity)

elif page == "totalTradedValue":
    st.title("totalTradedValue")
    df_totalTradedValue = pd.read_excel(excel_file, sheet_name="totalTradedValue")
    st.write(df_totalTradedValue)

elif page == "marketCapFreeFloat":
    st.title("marketCapFreeFloat")
    df_marketCapFreeFloat = pd.read_excel(excel_file, sheet_name="marketCapFreeFloat")
    st.write(df_marketCapFreeFloat)

elif page == "marketCapFull":
    st.title("marketCapFull")
    df_marketCapFull = pd.read_excel(excel_file, sheet_name="marketCapFull")
    st.write(df_marketCapFull)
