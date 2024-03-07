import quandl

# Set your Quandl API key
quandl.ApiConfig.api_key = "YOUR_API_KEY"

# Specify the symbol or scrip code for the stock you want to retrieve data for
scrip_code = "BSE/BOM500325"  # Example: 'BSE/BOM500325' for Reliance Industries on BSE

# Retrieve the stock data from Quandl
stock_data = quandl.get(scrip_code)

# Print the retrieved stock data
print(stock_data)
