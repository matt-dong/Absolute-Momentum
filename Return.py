# -*- coding: utf-8 -*-
"""
Created on Tue Jun 26 11:01:07 2018

@author: Matthew Dong
"""

import pandas as pd
import math
from pandas import ExcelWriter
from openpyxl import load_workbook

### load data from Excel
excel_file = 'PythonExport.xlsx'
start_date = pd.Timestamp("1973-01-31")
end_date = pd.Timestamp("2012-12-31")


### load tbill data to calculate sharpe
tbill = pd.read_excel('Absolute Momentum.xlsx', sheet_name = 'Monthly Returns')
tbill = tbill.set_index(['Date'])
tbill = tbill.loc[start_date:end_date]

info = []

### Desired asset class
### Change to 'M Monthly Returns' to use momentum data, also look at line 76 and make changes accordingly
asset = pd.read_excel(excel_file , sheet_name = 'Sheet1')
asset = asset.set_index(['Date'])
asset = asset.loc[start_date:end_date]
assetnames = list(asset.columns.values)

Tbill_Mean = (((tbill["TBILL"].mean()/100 + 1)**12 - 1)*100)

for assetname in assetnames:
    ### 1: Average Annualized Return
    Mean = (((asset[assetname].mean()/100 + 1)**12 - 1)*100)
    
    ### 2: Annualized Standard Deviation
    StdDev = ((asset[assetname].std() * math.sqrt(12)))
    
    ### 3: Sharpe Ratio
    Sharpe_Ratio = (Mean - Tbill_Mean)/StdDev
    
    ### 4: Percentage of Profit Months
    counter = 0
    for x in range(0, len(asset.index)):
        if asset[assetname].iloc[x] > 0:
            counter = counter + 1
    Profit_Months = (100 * counter / len(asset.index))
    
    ### 5: Maximum Drawdown
    investment = 100
    price = pd.DataFrame([100], columns=list('A'))
    for x in range(0, len(asset.index)):
        investment = investment * ((asset[assetname].iloc[x] / 100) + 1)
        price = price.append({"A": investment}, ignore_index=True)
    window = len(asset.index)
    Roll_Max = price["A"].rolling(window, min_periods=1).max()
    Monthly_Drawdown = price["A"]/Roll_Max - 1.0
    Max_Monthly_Drawdown = Monthly_Drawdown.rolling(window, min_periods=1).min()
    
    ### Add calculated results to a list
    info.append([str(assetname), Mean, StdDev, Sharpe_Ratio, Profit_Months, float(Max_Monthly_Drawdown.min()*100)])


### create dataframe containing results
data = pd.DataFrame(info,columns=['Ticker','Average Annualized Return', 'Annualized Standard Deviation', 'Sharpe Ratio', 'Profit Months', 'Max Drawdown'])
data = data.set_index(['Ticker'])

### Save data to excel workbook
book = load_workbook('PythonExport.xlsx')
writer = ExcelWriter('PythonExport.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
### Change to 'M Results' to save as momentum data
data.to_excel(writer, "M Index Results")
writer.save()   
