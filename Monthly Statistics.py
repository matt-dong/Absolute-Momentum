# -*- coding: utf-8 -*-
"""
Created on Mon Jun 25 13:59:30 2018

@author: Matthew Dong
"""

import pandas as pd
import numpy as np
from functools import reduce
import math
from pandas import ExcelWriter
from openpyxl import load_workbook


""" General Notes
TBILL results are spliced out because it is always last in the excel spreadsheet
Assume that the first investment of 100 is neither a buy nor a sell
"""

### TODO : MAKE SURE TO SET TIMEFRAME NAME BEFORE RUNNING
start_date = pd.Timestamp("1973-01-31")
end_date = pd.Timestamp("2012-12-31")

### load tbill data
excel_file = 'Absolute Momentum.xlsx'
tbill = pd.read_excel(excel_file, sheet_name = 'Monthly Returns')
tbill = tbill.set_index(['Date'])
tbill = tbill.loc[start_date:end_date]

### Import all asset data 
asset = pd.read_excel('Absolute Momentum.xlsx' , sheet_name = 'Monthly Returns')
asset = asset.set_index(['Date'])
assetnames = list(asset.columns.values)
assetnamesmomentum = [str(assetname) + " Momentum" for assetname in assetnames]

### create list to store data
month_info = []

### Loops momentum stuff for all assets
for y in range(0, len(assetnames)-1):
    
    ### Make seperate return dataframes
    asset = asset.loc[start_date:end_date]
    tbill_return = tbill[["TBILL"]]
    asset_return = asset[[assetnames[y]]]

    ### Create dataframes to store momentum
    tbill_momentum = pd.DataFrame(columns = ['Momentum'])
    asset_momentum = pd.DataFrame(columns = ['Momentum']) 
    
    ### Convert monthly returns into ratios so they can be multiplied
    tbill_return = (tbill_return/100)+1
    asset_return = (asset_return/100)+1
    
    ### momentum calculation, currently goes up to the latest date in the excel table
    period = 12
    for x in range(period, int(len(asset_return.index))):
        
        ### reduce((lambda... multiplies the 12 previous months' ratios and calculates the momentum
        asset_momentum = asset_momentum.append({'Momentum': reduce((lambda x,y: x*y), asset_return.values[int(x-period):int(x)])[0]}, ignore_index=True)   
        tbill_momentum = tbill_momentum.append({'Momentum': reduce((lambda x,y: x*y), tbill_return.values[int(x-period):int(x)])[0]}, ignore_index=True)
        
    ### set indexes for the two dataframes for less confusion
    tbill_momentum['Date'] = asset.index.values[period:len(asset.index)]
    tbill_momentum = tbill_momentum.set_index(['Date'])
    asset_momentum['Date'] = asset.index.values[period:len(asset.index)]
    asset_momentum = asset_momentum.set_index(['Date'])
    
    ### make sure returns and momentums match, hardcoding is removed
    signals = asset_momentum.subtract(tbill_momentum, fill_value=0)
    asset_return = asset_return[period:len(asset_return.index)]
    tbill_return = tbill_return[period:len(tbill_return.index)]
    
    ### Create signals
    signals['signal'] = 0.0
    signals['signal'] = np.where(signals['Momentum'] > 0 , 1.0, 0.0)
    
    ### Generate trading orders
    signals['positions'] = signals['signal'].diff()
       
    ### create dataframe for price of asset using momentum
    assetnamesmomentum[y] = pd.DataFrame([100], columns=list('A'))
    assetnamesmomentum[y] = assetnamesmomentum[y].rename(index=str, columns={"A": assetnames[y] + " Momentum"})
    investment = 100
    
    ### create variables to count months spent in asset/tbill and number of buy/sell signals
    asset_months = 0
    tbill_months = 0
    buy_signals = 0
    sell_signals = 0      
            
    ### momentum implementation
    for x in range(0, int(len(signals.index))):
    
        ### momentum is positive, we long the asset
        if signals.iloc[x]['signal'] == 1:
            a = x
            asset_months = asset_months + 1
            
            ### asset is already longed
            if signals.iloc[a]['positions'] == 0:
                investment = investment * asset_return.iloc[a][assetnames[y]]
                assetnamesmomentum[y] = assetnamesmomentum[y].append({assetnames[y] + " Momentum": investment}, ignore_index=True)
            
            ### asset is moving from tbill to long position, 20 point penalty
            elif signals.iloc[a]['positions'] == 1:
                buy_signals = buy_signals + 1
                investment = investment * 0.998 * asset_return.iloc[a][assetnames[y]]
                assetnamesmomentum[y] = assetnamesmomentum[y].append({assetnames[y] + " Momentum": investment}, ignore_index=True)
    
        ### momentum is negative, use tbill return
        elif signals.iloc[x]['signal'] == 0:
            a = x
            tbill_months = tbill_months + 1
            
            ### already in tbill
            if signals.iloc[a]['positions'] == 0:
                investment = investment * tbill_return.iloc[a]['TBILL']
                assetnamesmomentum[y] = assetnamesmomentum[y].append({assetnames[y] + " Momentum": investment}, ignore_index=True)
            
            ### asset is moving from long to tbill, 20 point penalty
            elif signals.iloc[a]['positions'] == -1:
                sell_signals = sell_signals + 1
                investment = investment * 0.998 * tbill_return.iloc[a]['TBILL']
                assetnamesmomentum[y] = assetnamesmomentum[y].append({assetnames[y] + " Momentum": investment}, ignore_index=True)
    

    
    ### check to see that the loop is working and append data to list
    print(assetnames[y])
    month_info.append([assetnames[y], asset_months, tbill_months, buy_signals, sell_signals, int(asset_months)/int(buy_signals), int(tbill_months)/int(sell_signals)])
    
    ### set index for prices to reduce confusion
    assetnamesmomentum[y]['Date'] = asset.index.values[period:len(asset.index)]
    assetnamesmomentum[y] = assetnamesmomentum[y].set_index(['Date'])

### create dataframe containing month info
month_data = pd.DataFrame(month_info,columns=['Ticker','Months in Asset', 'Months in TBill', 'Number of Buy Signals', 'Number of Sell Signals', "Average Duration in Asset", "Average Duration in TBill"])
month_data = month_data.set_index(['Ticker'])
print(month_data.to_string())

### merge the individual price dataframes into one

Momentum_Prices = tbill_momentum
del Momentum_Prices["Momentum"]
for x in range(0, len(assetnames)-1):
    Momentum_Prices = pd.concat([Momentum_Prices, assetnamesmomentum[x]], axis=1)

### save Momentum results to a spreadsheet
book = load_workbook('PythonExport.xlsx')
writer = ExcelWriter('PythonExport.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
Momentum_Prices.to_excel(writer, "Momentum Indexes")
###month_data.to_excel(writer, "Month Data")
writer.save()   
