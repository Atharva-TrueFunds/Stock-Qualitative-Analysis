import numpy as np
import pandas as pd
from bsedata.bse import BSE

data = pd.read_csv("Merge_currentValue.csv")

b = BSE()
data.rename(columns={'B': 'day1', 'C': 'day2', 'D': 'day3', 'E': 'day4', 'F': 'day5'}, inplace=True)
currentValue_data = data[['companyName', 'SecurityCode', 'day1', 'day2', 'day3', 'day4', 'day5']]

highest_column_name = []
highest_value = []
lowest_column_name = []
lowest_value = []

currentValue_data[['day1', 'day2', 'day3', 'day4', 'day5']] = currentValue_data[['day1', 'day2', 'day3', 'day4', 'day5']].apply(pd.to_numeric, errors='coerce')

for index, row in currentValue_data.iterrows():
    highest_value.append(row[['day1', 'day2', 'day3', 'day4', 'day5']].max())
    highest_column_name.append(row[['day1', 'day2', 'day3', 'day4', 'day5']].idxmax())
    lowest_value.append(row[['day1', 'day2', 'day3', 'day4', 'day5']].min())
    lowest_column_name.append(row[['day1', 'day2', 'day3', 'day4', 'day5']].idxmin())

for index, row in currentValue_data.iterrows():
    if len(set(row[['day1', 'day2', 'day3', 'day4', 'day5']])) > 1:
        data['Highest_Value'] = highest_value
        data['Highest_Column_Name'] = highest_column_name
        data['Lowest_Value'] = lowest_value
        data['Lowest_Column_Name'] = lowest_column_name

ascending_order = (data[['day1', 'day2', 'day3', 'day4', 'day5']].values ==
                    np.sort(data[['day1', 'day2', 'day3', 'day4', 'day5']], axis=1)).all(axis=1)

descending_order = (data[['day1', 'day2', 'day3', 'day4', 'day5']].values ==
                    np.sort(data[['day1', 'day2', 'day3', 'day4', 'day5']], axis=1)[:, ::-1]).all(axis=1)

ascending_data = data[ascending_order]
descending_data = data[descending_order]

print("Security Codes and Company Names with ascending order:")
print(ascending_data[['SecurityCode', 'companyName']])
print("\nSecurity Codes and Company Names with descending order:")
print(descending_data[['SecurityCode', 'companyName']])

data.to_csv('Merge_currentValue.csv', index=False)