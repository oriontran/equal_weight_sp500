# Equal Weight S&P500
## Mimics S&P500 index fund holdings with equal weighting
Uses a CSV file with tickers of the S&P500 to create and make batch requests 
from IEX API using sandbox token (for cost saving purposes) that provides
market data used to determine number of shares to purchase for each ticker
given a particular portfolio value. Creates and formats an XLSX file used to
save data and calculations into.
