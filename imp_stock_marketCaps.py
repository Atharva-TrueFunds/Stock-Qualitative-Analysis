import pandas as pd
from bsedata.bse import BSE

b = BSE()
data = pd.read_csv("marketcap_level.csv")

symbols = data['SecurityCode'].tolist()

extracted_data = []

for symbol in symbols:
    try:

        stock_quote = b.getQuote(str(symbol))

        if 'companyName' in stock_quote and 'marketCapFull' in stock_quote:

            extracted_info = {
                'SecurityCode': symbol,
                'companyName': stock_quote['companyName'],
                'marketCapFull': stock_quote['marketCapFull']
            }
            extracted_data.append(extracted_info)
    except Exception as e:
        pass

extracted_df = pd.DataFrame(extracted_data)

print(extracted_df)

extracted_df.to_csv('marketcap_level.csv', index=False)
