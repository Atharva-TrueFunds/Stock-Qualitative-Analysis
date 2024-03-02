import numpy as np
import pandas as pd
from bsedata.bse import BSE

data = pd.read_excel("Qualitative data analysis.xlsx", sheet_name="currentValue")

df = pd.DataFrame(data)

bse = BSE()
print(df)

symbols = data["scripCode"].tolist()


symbols_list = []
company_name_list = []
day_from_list = []
start_date_list = []
day_to_list = []
end_date_list = []
percent_change_list = []

up_lists_2_5 = []
up_lists_5_10 = []
up_lists_10_15 = []
up_lists_15_20 = []
up_lists_above_20 = []

up_lists = [[] for _ in range(21)]
down_lists = [[] for _ in range(21)]
# result=[]
for symbols in symbols:
    row_values = data[data["scripCode"] == symbols]
    if not row_values.empty:
        CompanyName = row_values["companyName"].values[0]
    else:
        print("DataFrame 'row_values' is empty.")

    for i in range(2, 6):
        for j in range(2, 6):
            if i != j:
                day_from = i
                day_to = j
                Start_Date = pd.to_numeric(
                    row_values.iloc[0, day_from], errors="coerce"
                )
                End_Date = pd.to_numeric(row_values.iloc[0, day_to], errors="coerce")
                percent_change = (((End_Date - Start_Date) / Start_Date) * 100).round(2)

                symbols_list.append(symbols)
                company_name_list.append(CompanyName)
                day_from_list.append(day_from - 1)
                start_date_list.append(Start_Date)
                day_to_list.append(day_to - 1)
                end_date_list.append(End_Date)
                percent_change_list.append(percent_change)

                df = pd.DataFrame(
                    {
                        "Symbol": symbols_list,
                        "Company Name": company_name_list,
                        "Day From": day_from_list,
                        "Start Date": start_date_list,
                        "Day To": day_to_list,
                        "End Date": end_date_list,
                        "Percent Change": percent_change_list,
                    }
                )

                greater_than_df = df[df["Start Date"] > df["End Date"]]

                smaller_than_df = df[df["Start Date"] < df["End Date"]]

                with pd.ExcelWriter("price_Change_day_to_each_day.xlsx") as writer:
                    greater_than_df.to_excel(writer, sheet_name="Data", index=False)
                    smaller_than_df.to_excel(
                        writer,
                        sheet_name="Data",
                        startrow=len(greater_than_df) + 2,
                        index=False,
                    )

                if percent_change >= 2 and percent_change <= 5:
                    up_lists_2_5.append(
                        (symbols, CompanyName, day_from - 1, day_to - 1, percent_change)
                    )
                elif percent_change > 5 and percent_change <= 10:
                    up_lists_5_10.append(
                        (symbols, CompanyName, day_from - 1, day_to - 1, percent_change)
                    )
                elif percent_change > 10 and percent_change <= 15:
                    up_lists_10_15.append(
                        (symbols, CompanyName, day_from - 1, day_to - 1, percent_change)
                    )
                elif percent_change > 15 and percent_change <= 20:
                    up_lists_15_20.append(
                        (symbols, CompanyName, day_from - 1, day_to - 1, percent_change)
                    )
                elif percent_change > 20:
                    up_lists_above_20.append(
                        (symbols, CompanyName, day_from - 1, day_to - 1, percent_change)
                    )

df_2_5 = pd.DataFrame(
    up_lists_2_5,
    columns=["Symbol", "Company Name", "Day From", "Day To", "Percent Change 2-5%"],
)
df_5_10 = pd.DataFrame(
    up_lists_5_10,
    columns=["Symbol", "Company Name", "Day From", "Day To", "Percent Change 5-10%"],
)
df_10_15 = pd.DataFrame(
    up_lists_10_15,
    columns=["Symbol", "Company Name", "Day From", "Day To", "Percent Change 10-15%"],
)
df_15_20 = pd.DataFrame(
    up_lists_15_20,
    columns=["Symbol", "Company Name", "Day From", "Day To", "Percent Change 15-20%"],
)
df_above_20 = pd.DataFrame(
    up_lists_above_20,
    columns=[
        "Symbol",
        "Company Name",
        "Day From",
        "Day To",
        "Percent Change above 20%",
    ],
)

with pd.ExcelWriter("price_Change_day_to_day.xlsx") as writer:
    df_2_5.to_excel(writer, sheet_name="2-5%", index=False)
    df_5_10.to_excel(writer, sheet_name="5-10%", index=False)
    df_10_15.to_excel(writer, sheet_name="10-15%", index=False)
    df_15_20.to_excel(writer, sheet_name="15-20%", index=False)
    df_above_20.to_excel(writer, sheet_name="Above 20%", index=False)
