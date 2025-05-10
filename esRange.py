import yfinance as yf
import pandas as pd
from datetime import datetime

# ticker and date range
ticker = 'ES=F'
start_date = '2025-01-01'
end_date = datetime.today().strftime('%Y-%m-%d')

# download ES futures data
data = yf.download(ticker, start=start_date, end=end_date, auto_adjust=False)

# show what Yahoo returned
print("\n Columns returned from Yahoo Finance:")
print(list(data.columns))

print("\ first rows:")
print(data.head())

# check if we can calculate range
if 'High' in data.columns and 'Low' in data.columns:
    # Drop missing data and compute range
    clean_data = data.dropna(subset=['High', 'Low'])
    clean_data['ES_Range'] = clean_data['High'] - clean_data['Low']
    
    # print final table
    result = clean_data[['High', 'Low', 'ES_Range']].copy()
    result.columns = ['ES_High', 'ES_Low', 'ES_Range']
    result.index.name = 'Date'

    print("\n Final ES range data:")
    print(result)

else:
    #
    print("\n'High' and 'Low' columns not available")
    print(":white_check_mark: Available columns:", list(data.columns))
