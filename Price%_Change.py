import numpy as np
import pandas as pd
from bsedata.bse import BSE

data = pd.read_excel("Price_Change.xlsx", sheet_name='price_change')
bse = BSE()

symbols = data['SecurityCode'].tolist()

up_lists = [[] for _ in range(21)]
down_lists = [[] for _ in range(21)]

for symbols in symbols:
    row_values = data[data['SecurityCode'] == symbols]
    CompanyName= row_values['companyName'].values[0]
    for i in range (1,11):
        for j in range(1,11):
            if i!=j:
                day_from = i
                day_to = j
                Start_Date = row_values[day_from].values[0]
                End_Date = row_values[day_to].values[0]
                percent_change = (((End_Date - Start_Date) / Start_Date) * 100).round(2)
                print(f"The Percent Change for {symbols} {CompanyName} from Day{day_from} '{Start_Date}' to Day{day_to} '{End_Date}' is {percent_change}% ")

                if percent_change >= 2 and percent_change <= 5:
                    up_lists[2].append(f"Stock Up more than 2% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")
                elif percent_change <= -2 and percent_change >= -5:
                    down_lists[2].append(f"Stock Down more than 2% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")
                elif percent_change >= 5 and percent_change <= 10:
                    up_lists[5].append(f"Stock Up more than 5% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")
                elif percent_change <= -5 and percent_change >= -10:
                    down_lists[5].append(f"Stock Down more than 5% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")
                elif percent_change >= 10 and percent_change <= 15:
                    up_lists[10].append(f"Stock Up more than 10% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")
                elif percent_change <= -10 and percent_change >= -15:
                    down_lists[10].append(f"Stock Down more than 10% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")
                elif percent_change >= 15 and percent_change <= 20:
                    up_lists[15].append(f"Stock Up more than 15% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")
                elif percent_change <= -15 and percent_change >= -20:
                    down_lists[15].append(f"Stock Down more than 15% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")
                elif percent_change >= 20:
                    up_lists[20].append(f"Stock Up more than 20% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")
                elif percent_change <= -20:
                    down_lists[20].append(f"Stock Down more than 20% is {CompanyName} from Day{day_from} to Day{day_to} which is {percent_change}%")

print(f"\nUp more than {2}%: ")
print("\n".join(up_lists[2]))
print(f"\nDown more than {2}%: ")
print("\n".join(down_lists[2]))

print(f"\nUp more than {5}%: ")
print("\n".join(up_lists[5]))
print(f"\nDown more than {5}%: ")
print("\n".join(down_lists[5]))

print(f"\nUp more than {10}%: ")
print("\n".join(up_lists[10]))
print(f"\nDown more than {10}%: ")
print("\n".join(down_lists[10]))

print(f"\nUp more than {15}%: ")
print("\n".join(up_lists[15]))
print(f"\nDown more than {15}%: ")
print("\n".join(down_lists[15]))

print(f"\nUp more than {20}%: ")
print("\n".join(up_lists[20]))
print(f"\nDown more than {20}%: ")
print("\n".join(down_lists[20]))
