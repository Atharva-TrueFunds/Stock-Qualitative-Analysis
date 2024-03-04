import pandas as pd
import streamlit as st

data = pd.read_excel("All_sheets.xlsx", sheet_name="currentValue")
print(data.columns)
data.columns.values[12:13] = ["04Mar"]
data["percent_diff_up"] = ((data["52weekHigh"] - data["04Mar"]) / data["04Mar"]) * 100
filtered_data_0_2 = data[(data["percent_diff_up"] >= 0) & (data["percent_diff_up"] < 2)]
filtered_data_2_5 = data[(data["percent_diff_up"] >= 2) & (data["percent_diff_up"] < 5)]
filtered_data_5_10 = data[
    (data["percent_diff_up"] >= 5) & (data["percent_diff_up"] <= 10)
]
filtered_data_10_15 = data[
    (data["percent_diff_up"] >= 10) & (data["percent_diff_up"] < 15)
]
filtered_data_15_20 = data[
    (data["percent_diff_up"] >= 15) & (data["percent_diff_up"] < 20)
]
filtered_data_20 = data[(data["percent_diff_up"] >= 20)]


df_up_0 = filtered_data_0_2[
    ["scripCode", "companyName", "04Mar", "52weekHigh", "percent_diff_up"]
]

df_up_2 = filtered_data_2_5[
    ["scripCode", "companyName", "04Mar", "52weekHigh", "percent_diff_up"]
]

df_up_5 = filtered_data_5_10[
    ["scripCode", "companyName", "04Mar", "52weekHigh", "percent_diff_up"]
]

df_up_10 = filtered_data_10_15[
    ["scripCode", "companyName", "04Mar", "52weekHigh", "percent_diff_up"]
]

df_up_15 = filtered_data_15_20[
    ["scripCode", "companyName", "04Mar", "52weekHigh", "percent_diff_up"]
]
# df_up_20 = filtered_data_20[
#     ["scripCode", "companyName", "04Mar", "52weekHigh", "percent_diff_up"]
# ]
# st.table(df_up_20)


data["percent_diff_up"] = ((data["52weekLow"] - data["04Mar"]) / data["04Mar"]) * 100
filtered_data_0_2 = data[(data["percent_diff_up"] >= 0) & (data["percent_diff_up"] < 2)]
filtered_data_2_5 = data[(data["percent_diff_up"] >= 2) & (data["percent_diff_up"] < 5)]
filtered_data_5_10 = data[
    (data["percent_diff_up"] >= 5) & (data["percent_diff_up"] <= 10)
]
filtered_data_10_15 = data[
    (data["percent_diff_up"] >= 10) & (data["percent_diff_up"] < 15)
]
filtered_data_15_20 = data[
    (data["percent_diff_up"] >= 15) & (data["percent_diff_up"] < 20)
]
filtered_data_20 = data[(data["percent_diff_up"] >= 20)]

df_down_0 = filtered_data_0_2[
    ["scripCode", "companyName", "04Mar", "52weekLow", "percent_diff_up"]
]

df_down_2 = filtered_data_2_5[
    ["scripCode", "companyName", "04Mar", "52weekLow", "percent_diff_up"]
]

df_down_5 = filtered_data_5_10[
    ["scripCode", "companyName", "04Mar", "52weekLow", "percent_diff_up"]
]

df_down_10 = filtered_data_10_15[
    ["scripCode", "companyName", "04Mar", "52weekLow", "percent_diff_up"]
]

df_down_15 = filtered_data_15_20[
    ["scripCode", "companyName", "04Mar", "52weekLow", "percent_diff_up"]
]

st.write("Stock 0% to 2% UP diffenrce")
st.table(df_up_0)
st.write("Stock 0% to 2% DOWN diffenrce")
st.table(df_down_0)

st.write("Stock 2% to 5% UP diffenrce")
st.table(df_up_2)
st.write("Stock 2% to 5% Down diffenrce")
st.table(df_down_2)

st.write("Stock 5% to 10% UP diffenrce")
st.table(df_up_5)
st.write("Stock 5% to 10% DOWN diffenrce")
st.table(df_down_5)

st.write("Stock 10% to 15% UP diffenrce")
st.table(df_up_10)
st.write("Stock 10% to 15% Down diffenrce")
st.table(df_down_10)

st.write("Stock 15% to 20% UP diffenrce")
st.table(df_up_15)
st.write("Stock 15% to 20% DOWN diffenrce")
st.table(df_down_15)


# df_down_20 = filtered_data_20[
#     ["scripCode", "companyName", "04Mar", "52weekHigh", "percent_diff_up"]
# ]
# st.table(df_down_20)

# with pd.ExcelWriter("near_52week.xlsx", engine="xlsxwriter") as writer:

#     df_up_0.to_excel(writer, sheet_name="0-2%", index=True, startrow=0, startcol=0)
#     df_down_0.to_excel(
#         writer, sheet_name="0-2%", index=True, startrow=df_up_0.shape[0] + 2, startcol=0
#     )

#     df_up_2.to_excel(writer, sheet_name="2-5%", index=True, startrow=0, startcol=0)
#     df_down_2.to_excel(
#         writer, sheet_name="2-5%", index=True, startrow=df_up_2.shape[0] + 2, startcol=0
#     )

#     df_up_5.to_excel(writer, sheet_name="5-10%", index=True, startrow=0, startcol=0)
#     df_down_5.to_excel(
#         writer,
#         sheet_name="5-10%",
#         index=True,
#         startrow=df_up_5.shape[0] + 2,
#         startcol=0,
#     )

#     df_up_10.to_excel(writer, sheet_name="10-15%", index=True, startrow=0, startcol=0)
#     df_down_10.to_excel(
#         writer,
#         sheet_name="10-15%",
#         index=True,
#         startrow=df_up_10.shape[0] + 2,
#         startcol=0,
#     )

#     df_up_15.to_excel(writer, sheet_name="15-20%", index=True, startrow=0, startcol=0)
#     df_down_15.to_excel(
#         writer,
#         sheet_name="15-20%",
#         index=True,
#         startrow=df_up_15.shape[0] + 2,
#         startcol=0,
#     )

#     df_up_20.to_excel(writer, sheet_name="+20%", index=True, startrow=0, startcol=0)
#     df_down_20.to_excel(
#         writer,
#         sheet_name="+20%",
#         index=True,
#         startrow=df_up_20.shape[0] + 2,
#         startcol=0,
#     )
