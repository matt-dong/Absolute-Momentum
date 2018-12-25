# -*- coding: utf-8 -*-
"""
Created on Thu Jul 26 10:16:09 2018

@author: Matthew Dong
"""
import pandas as pd
import urllib.request
import json
from pandas import ExcelWriter
import time
from openpyxl import load_workbook


### TODO : MAKE SURE TO SET TIMEFRAME NAME BEFORE RUNNING
start_date = pd.Timestamp("1990-02-28")
end_date = pd.Timestamp("2018-04-30")

### load tbill data for sharpe calculation
excel_file = 'Absolute Momentum.xlsx'
tbill = pd.read_excel(excel_file, sheet_name = 'Monthly Returns')
tbill = tbill.set_index(['Date'])
tbill = tbill.loc[start_date:end_date]
tbill_return = tbill[["TBILL"]]

### create dataframe to store all prices
prices = tbill_return
del prices["TBILL"]


# download data
tickers = ["MMM", "AXP", "AAPL", "BA", "CAT", "CVX", "KO", "DIS", "XOM", "HD", "IBM"]
tickernames = tickers
for a in range(0, len(tickers)):
    x = str(tickers[a])
    
    ### Alpha Vantage has a limit of 3 calls/minute, so I have added a 60 second delay on every 3rd ticker
    if a % 3 == 2: 
        with urllib.request.urlopen("https://www.alphavantage.co/query?function=TIME_SERIES_MONTHLY_ADJUSTED&symbol=" + tickers[a] + "&apikey=2QQ1TJQIHMI4DBOH") as url:
            x = json.loads(url.read().decode())
            
            ### Use Adjusted Monthly Data
            x = pd.DataFrame.from_dict(x["Monthly Adjusted Time Series"])
            x = x.transpose()
            x = x.iloc[::-1]
            x = x.rename(index=str, columns={"5. adjusted close": str(tickers[a])})
            x = x.loc[:str(end_date)]
            x['Date'] = tbill_return.index.values
            x = x.set_index(['Date'])
            prices = pd.concat([prices, x[str(tickers[a])]], axis=1)
            print(str(x))
            time.sleep(60)
    
    ### same code as above
    else:
        with urllib.request.urlopen("https://www.alphavantage.co/query?function=TIME_SERIES_MONTHLY_ADJUSTED&symbol=" + tickers[a] + "&apikey=2QQ1TJQIHMI4DBOH") as url:
            x = json.loads(url.read().decode())
            x = pd.DataFrame.from_dict(x["Monthly Adjusted Time Series"])
            x = x.transpose()
            x = x.iloc[::-1]
            x = x.rename(index=str, columns={"5. adjusted close": str(tickers[a])})
            x = x.loc[:str(end_date)]
            x['Date'] = tbill_return.index.values
            x = x.set_index(['Date'])
            prices = pd.concat([prices, x[str(tickers[a])]], axis=1)
            print(str(x))
            
### saves pricing data to a python excel file
book = load_workbook('PythonExport.xlsx')
writer = ExcelWriter('PythonExport.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
prices.to_excel(writer, "Prices")
writer.save()