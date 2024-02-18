import numpy as np
import pandas as pd
from bsedata.bse import BSE

data= pd.read_csv("Merge_currentValue.csv")

data.rename(columns={'B': 'day1', 'C': 'day2', 'D': 'day3', 'E': 'day4', 'F': 'day5'}, inplace=True)
currentValue_data = data[['day1', 'day2', 'day3', 'day4', 'day5']]
b=BSE()

highest_column_name=[]
highest_value=[]
lowest_column_name=[]
lowest_value=[]

for index, row in currentValue_data.iterrows():
    highest_value.append(row.max())
    highest_column_name.append(row.idxmax())
    lowest_value.append(row.min())
    lowest_column_name.append(row.idxmin())

data['Highest_Value'] = highest_value
data['Highest_Column_Name'] = highest_column_name
data['Lowest_Value'] = lowest_value
data['Lowest_Column_Name'] = lowest_column_name

data.to_csv('Merge_currentValue.csv', index=False)