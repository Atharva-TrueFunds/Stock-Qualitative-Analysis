import numpy as np
import pandas as pd

data = pd.read_csv('marketcap_level.csv')

last_column_index = len(data.columns) - 1

next_column_name = int(ord('A') + last_column_index)

data['marketCapFull'] = data['marketCapFull'].replace({'\$': '', ',': '', ' Cr.': ''}, regex=True).astype(float)

descending_data = data.sort_values(by='marketCapFull')[::-1]

print(descending_data)

descending_data.to_csv('marketcap_level.csv', index=False)
