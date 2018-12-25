# -*- coding: utf-8 -*-
"""
Created on Mon Jun 25 13:59:30 2018

@author: Matthew Dong
"""

import pandas as pd



### load data
excel_file = 'Absolute Momentum.xlsx'


EAFE = pd.read_excel(excel_file, sheet_name = 'MSCI')
EAFE = EAFE.set_index(['Date'])
EAFE_return= EAFE[["EAFE Return"]][12:480]

USA = pd.read_excel(excel_file, sheet_name = 'MSCI')
USA = USA.set_index(['Date'])
USA_return = USA[["USA Return"]][12:480]

TBOND = pd.read_excel(excel_file, sheet_name = 'Treasury')
TBOND = TBOND.set_index(['Date'])
TBOND_return= TBOND[["TBOND Return"]][12:480]


Credit = pd.read_excel(excel_file, sheet_name = 'US Government&Credit')
Credit = Credit.set_index(['Date'])
Credit_return= Credit[["Credit Return"]][12:480]

REIT = pd.read_excel(excel_file, sheet_name = 'REIT')
REIT = REIT.set_index(['Date'])
REIT_return = REIT[["REIT Return"]][12:480]

GSCI = pd.read_excel(excel_file, sheet_name = 'GSCI')
GSCI = GSCI.set_index(['Date'])
GSCI_return = GSCI[["GSCI Return"]][12:480]

Gold = pd.read_excel(excel_file, sheet_name = 'Gold')
Gold = Gold.set_index(['Date'])
Gold_return = Gold[["Gold Return"]][12:480]

HiYld = pd.read_excel(excel_file, sheet_name = 'US Corporate High Yield')
HiYld = HiYld.set_index(['Date'])
HiYld_return = HiYld[["HiYld Return"]][12:480]

assets = [EAFE_return, USA_return, TBOND_return, Credit_return, HiYld_return, REIT_return, GSCI_return, Gold_return]
assetnames = ["EAFE Return", "USA Return", "TBOND Return", "Credit Return",  "HiYld Return", "REIT Return", "GSCI Return", "Gold Return"]
a = assets[1:]
anames = assetnames[1:]

for x in range(0,7):
    print(str(anames[x]) + " " + str(EAFE_return["EAFE Return"].corr(a[x][anames[x]])))

