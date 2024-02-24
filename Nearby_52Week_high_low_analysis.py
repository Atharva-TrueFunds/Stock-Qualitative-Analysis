import pandas as pd

data = pd.read_excel("Qualitative data analysis.xlsx", sheet_name='demo')

data['52weekHigh'] = pd.to_numeric(data['52weekHigh'], errors='coerce').round(2)
data['52weekLow'] = pd.to_numeric(data['52weekLow'], errors='coerce').round(2)

imp_columns_up = data[['scripCode', 'companyName', 'currentValue', '52weekHigh']]
imp_columns_down = data[['scripCode', 'companyName', 'currentValue', '52weekLow']]

imp_columns_up['10%CurrentValue'] = imp_columns_up['currentValue'] * 0.10
imp_columns_down['10%CurrentValue'] = imp_columns_down['currentValue'] * 0.10

imp_columns_up['UpperThreshold'] = imp_columns_up['currentValue'] + imp_columns_up['10%CurrentValue']
imp_columns_down['LowerThreshold'] = imp_columns_down['currentValue'] - imp_columns_down['10%CurrentValue']

upperThreshold_list = [int(value) for value in imp_columns_up['UpperThreshold']]
lowerThreshold_list = [int(value) for value in imp_columns_down['LowerThreshold']]

up = []
down = []
for i in range(len(upperThreshold_list)):
    if (upperThreshold_list[i] >= 0.05) and (imp_columns_up['52weekHigh'].iloc[i] <= 0.11):
        up.append((imp_columns_up['scripCode'].iloc[i], imp_columns_up['companyName'].iloc[i],
                   imp_columns_up['currentValue'].iloc[i], imp_columns_up['52weekHigh'].iloc[i]))

for i in range(len(lowerThreshold_list)):
    if (lowerThreshold_list[i] >= 0.05) and (imp_columns_down['52weekLow'].iloc[i] <= 0.11):
        down.append((imp_columns_down['scripCode'].iloc[i], imp_columns_down['companyName'].iloc[i],
                     imp_columns_down['currentValue'].iloc[i], imp_columns_down['52weekLow'].iloc[i]))

df_10_per_up = pd.DataFrame(up, columns=('scripCode', 'companyName', 'currentValue', '52weekHigh'))
df_10_per_down = pd.DataFrame(down, columns=('scripCode', 'companyName', 'currentValue', '52weekLow'))

with pd.ExcelWriter('near_52week.xlsx') as writer:
    df_10_per_up.to_excel(writer, sheet_name='10% Up', index=False)
    df_10_per_down.to_excel(writer, sheet_name='10% Down', index=False)
