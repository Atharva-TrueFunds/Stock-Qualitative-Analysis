import pandas as pd
import streamlit as st

data = pd.read_excel("All_sheets.xlsx", sheet_name="master")

data.columns.values[3:13] = [
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

imp_columns = data[
    [
        "scripCode",
        "companyName",
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
        "52weekHigh",
        "52weekLow",
    ]
]

imp_columns["52weekHigh"] = pd.to_numeric(
    imp_columns["52weekHigh"], errors="coerce"
).round(2)

companies_exceeding_52week_high = pd.DataFrame(columns=imp_columns.columns)
companies_exceeding_52week_low = pd.DataFrame(columns=imp_columns.columns)

for column in [
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
]:

    exceeding_52week_high = imp_columns[
        imp_columns[column] >= imp_columns["52weekHigh"]
    ]
    companies_exceeding_52week_high = pd.concat(
        [companies_exceeding_52week_high, exceeding_52week_high], ignore_index=False
    )

    exceeding_52week_low = imp_columns[imp_columns[column] <= imp_columns["52weekLow"]]
    companies_exceeding_52week_low = pd.concat(
        [companies_exceeding_52week_low, exceeding_52week_low], ignore_index=False
    )

companies_exceeding_52week_high = companies_exceeding_52week_high.drop_duplicates()
companies_exceeding_52week_low = companies_exceeding_52week_low.drop_duplicates()

exceeding_52week_high_data = []
for index, row in companies_exceeding_52week_high.iterrows():
    exceed_date = row[
        [
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
    ].idxmax()

    exceed_value = row[
        [
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
    ].max()

    exceeding_52week_high_data.append(
        [
            row["scripCode"],
            row["companyName"],
            exceed_date,
            exceed_value,
            row["52weekHigh"],
        ]
    )

exceeding_52week_low_data = []
for index, row in companies_exceeding_52week_low.iterrows():
    exceed_date = row[
        [
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
    ].idxmin()

    exceed_value = row[
        [
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
    ].min()

    exceeding_52week_low_data.append(
        [
            row["scripCode"],
            row["companyName"],
            exceed_date,
            exceed_value,
            row["52weekLow"],
        ]
    )

exceeding_52week_high_df = pd.DataFrame(
    exceeding_52week_high_data,
    columns=[
        "Scrip Code",
        "Company Name",
        "Exceeded High Date",
        "Value on that Date",
        "52-Week High",
    ],
)
exceeding_52week_high_df["52-Week High"] = exceeding_52week_high_df[
    "52-Week High"
].round(2)

st.table(exceeding_52week_high_df)

exceeding_52week_low_df = pd.DataFrame(
    exceeding_52week_low_data,
    columns=[
        "Scrip Code",
        "Company Name",
        "Exceeded Low Date",
        "Value on that day",
        "52-Week Low",
    ],
)

exceeding_52week_low_df["52-Week Low"] = exceeding_52week_low_df["52-Week Low"].round(2)

st.table(exceeding_52week_low_df)
