import pandas as pd
import requests
import math
IEX_CLOUD_API_TOKEN = 'Tpk_059b97af715d417d9f49f50b51b1c448'


def split_tickers(list_tickers, n):
    chunks = []
    for j in range(0, len(list_tickers), n):
        yield list_tickers[j:j + 100]


# Create csv object using a csv file
stocks = pd.read_csv('S&P500_Holdings.csv')

# Create a data frame with the proper column names using a list and sending to data frame call
my_cols = ['Tickers', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
final_data_frame = pd.DataFrame(columns=my_cols)

# Generator for list of stocks split into chunks of 100. Uses symbol column from csv object
ticker_chunks = split_tickers(stocks['Symbol'], 100)

# For each chunk of 100
for chunk in ticker_chunks:
    # create comma separated string from chunk
    symbol_string = ','.join(chunk)

    # Create batch api url using API documentation (uses symbol_string and token)
    batch_api_call = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}' \
                     f'&types=quote&token={IEX_CLOUD_API_TOKEN}'

    # Create python dictionary of JSON object created from the request response
    data = requests.get(batch_api_call).json()

    # For each symbol in the chunk of 100
    for symbol in chunk:
        # Pull the market cap and price for the symbol
        market_cap = data[symbol]['quote']['marketCap']
        price = data[symbol]['quote']['latestPrice']

        # Append the data into the data frame
        final_data_frame = final_data_frame.append(
            pd.Series(
                [
                    symbol, price, market_cap, 'N/A'
                ],
                index=my_cols
            ),
            ignore_index=True
        )

# Acquire size of portfolio
while True:
    try:
        portfolio_size = float(input("What is the value of your portfolio: "))
        break
    except ValueError:
        print("Invalid number. Please try again.")

# Determine size of position and corresponding number of shares for each then set row entry to that value
for i in range(len(final_data_frame)):
    position_size = portfolio_size/len(final_data_frame)
    shares = math.floor(position_size/final_data_frame.loc[i, "Stock Price"])
    final_data_frame.loc[i, "Number of Shares to Buy"] = shares

# Load data into xlsx file
writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
final_data_frame.to_excel(writer, "Recommended Trades", index=False)

# Determine colors and create format templates for xlsx writer
background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

int_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

# Apply formatting to each excel column
formatting = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Capitalization', dollar_format],
    'D': ['Number of Shares to Buy', int_format]
}

for key in formatting.keys():
    # Apply formatting to column
    writer.sheets['Recommended Trades'].set_column(f'{key}:{key}', 18, formatting[key][1])

    # Overwrite the column title
    writer.sheets['Recommended Trades'].write(f'{key}1', formatting[key][0], formatting[key][1])

writer.save()
