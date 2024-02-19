import numpy as np
import pandas as pd

data = pd.read_csv('Updated_stock_data_Active.csv')

data['marketCapFull'] = data['marketCapFull'].replace({'\$': '', ',': '', ' Cr.': ''}, regex=True).astype(float)

descending_data = data.sort_values(by='marketCapFull')[::-1]

print(descending_data)

descending_data.to_csv('Updated_stock_data_Active.csv', index=False)
