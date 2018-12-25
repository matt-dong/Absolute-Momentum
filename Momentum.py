# -*- coding: utf-8 -*-
"""
Created on Mon Jun 25 13:59:30 2018

@author: Matthew Dong
"""

import pandas as pd
import numpy as np



### load data
excel_file = 'Main Worksheet.xlsm'
tbill = pd.read_excel(excel_file, sheet_name = 'Treasury')
tbill = tbill.set_index(['Date'])
tbill_return= tbill[["tbill Return"]]

### Choose Asset Class
asset = pd.read_excel(excel_file , sheet_name = 'Gold')
asset = asset.set_index(['Date'])
asset_return = asset[["Return"]]
monthly_close = asset[["Momentum"]]


### momentum calculation
period = 12
asset_period_return = monthly_close / monthly_close.shift(period)

tbill_period_return = tbill[["Momentum"]]

signals = asset_period_return.subtract(tbill_period_return, fill_value=0)
signals = signals[12:480]
tbill_return = tbill_return[12:480]
asset_return = asset_return[12:480]

### Create signals
signals['signal'] = 0.0
signals['signal'] = np.where(signals['Momentum'] > 0 , 1.0, 0.0)

### Generate trading orders
signals['positions'] = signals['signal'].diff()

### create dataframe for prices
price = pd.DataFrame([100], columns=list('A'))
price = price.rename(index=str, columns={"A": "Price"})
 
investment = 100

### momentum implementation
for x in range(0,468):
    
    ### momentum is positive, we long the asset
    if signals.iloc[x]['signal'] == 1:
        a = x
        
        ### asset is already longed
        if signals.iloc[a]['positions'] == 0:
            investment = investment * asset_return.iloc[a]['Return']
            price = price.append({'Price': investment}, ignore_index=True)
            
        ### asset is moving from tbill to long position, 20 point penalty
        elif signals.iloc[a]['positions'] == 1:
            investment = investment * 0.998 * asset_return.iloc[a]['Return']
            price = price.append({'Price': investment}, ignore_index=True)
    
    ### momentum is negative, use tbill return
    elif signals.iloc[x]['signal'] == 0:
        a = x
        
        ### already in tbill
        if signals.iloc[a]['positions'] == 0:
            investment = investment * tbill_return.iloc[a]['tbill Return']
            price = price.append({'Price': investment}, ignore_index=True)
            
        ### asset is moving from long to tbill, 20 point penalty
        elif signals.iloc[a]['positions'] == -1:
            investment = investment * 0.998 * tbill_return.iloc[a]['tbill Return']
            price = price.append({'Price': investment}, ignore_index=True)
            
price['Date'] = asset.index.values[12:len(asset.index)]
price = price.set_index(['Date'])
print(price.to_string())
