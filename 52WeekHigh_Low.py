import numpy as np
import pandas as pd
from tabulate import tabulate

data = pd.read_excel("Qualitative data analysis.xlsx", sheet_name='currentValue')

data.columns.values[2:6] = ['19Feb', '20Feb', '21Feb', '22Feb']

imp_columns = data[['scripCode', 'companyName', '19Feb', '20Feb', '21Feb', '22Feb', '52weekHigh', '52weekLow']]

imp_columns['52weekHigh'] = pd.to_numeric(imp_columns['52weekHigh'], errors='coerce').round(2)

companies_exceeding_52week_high = pd.DataFrame(columns=imp_columns.columns)
companies_exceeding_52week_low = pd.DataFrame(columns=imp_columns.columns)

for column in ['19Feb', '20Feb', '21Feb', '22Feb']:

    exceeding_52week_high = imp_columns[imp_columns[column] >= imp_columns['52weekHigh']]
    companies_exceeding_52week_high = pd.concat([companies_exceeding_52week_high, exceeding_52week_high], ignore_index=False)

    exceeding_52week_low = imp_columns[imp_columns[column] <= imp_columns['52weekLow']]
    companies_exceeding_52week_low = pd.concat([companies_exceeding_52week_low, exceeding_52week_low], ignore_index=False)

companies_exceeding_52week_high = companies_exceeding_52week_high.drop_duplicates()
companies_exceeding_52week_low = companies_exceeding_52week_low.drop_duplicates()

exceeding_52week_high_data = []
for index, row in companies_exceeding_52week_high.iterrows():
    exceed_date = row[['19Feb', '20Feb', '21Feb', '22Feb']].idxmax()
    abc=exceeding_52week_high_data.append([row['scripCode'], row['companyName'], exceed_date, row['52weekHigh']])

exceeding_52week_low_data = []
for index, row in companies_exceeding_52week_low.iterrows():
    exceed_date = row[['19Feb', '20Feb', '21Feb', '22Feb']].idxmin()
    xyz=exceeding_52week_low_data.append([row['scripCode'], row['companyName'], exceed_date, row['52weekLow']])

exceeding_52week_high_df = pd.DataFrame(exceeding_52week_high_data, columns=['Scrip Code', 'Company Name', 'Exceeded High Date', '52-Week High'])

exceeding_52week_low_df = pd.DataFrame(exceeding_52week_low_data, columns=['Scrip Code', 'Company Name', 'Exceeded Low Date', '52-Week Low'])

import pandas as pd

with pd.ExcelWriter('52WeekHigh_Low.xlsx', engine='xlsxwriter') as writer:

    exceeding_52week_high_df.to_excel(writer, sheet_name='Exceeding 52_Week_high_low', index=True, startrow=0, startcol=0)

    workbook = writer.book
    worksheet = writer.sheets['Exceeding 52_Week_high_low']
    exceeding_52week_low_df.to_excel(writer, sheet_name='Exceeding 52_Week_high_low', index=True, startrow=exceeding_52week_high_df.shape[0] + 2, startcol=0)
