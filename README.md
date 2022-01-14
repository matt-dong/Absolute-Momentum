# Absolute-Momentum
A programming project coded in Python by Matthew Dong under the guidance of Oakland University Finance Professor Dr. Hong.

Implementation of Gary Antonnaci's absolute momentum trading strategy. His full report can be found at https://papers.ssrn.com/sol3/papers.cfm?abstract_id=2244633.

## Programs
[Get_Data.py](Get_Data.py) accesses the AlphaVantage API to retrieve historical stock data and stores the data in an excel workbook for future use.

[Correlation.py](Correlation.py) calculates the correlation between any two investments over a given period of time.

[Momentum.py](Momentum.py) generates trading orders and executes trades based on the historical pricing data for a single asset class. It also records the value of the portfolio over time and stores the data in an excel workbook for future use.

[Return.py](Return.py) calculates technical indicators such as Maximum Drawdown, Sharpe Ratio, Mean, Standard Deviation, and Percentage of profitable months for all asset classes. The data is compiled into a table, which is saved into an excel workbook for analysis and future use.

[Monthly Statistics.py](Monthly_Statistics.py) returns the results of the momentum trading. Number of months spent holding the asset, number of months spent holding treasury bonds, number of buy/sell signals, and the average duration (in months) spent in assets/treasury bonds before switching are compiled into a table and saved into an excel workbook for analysis and comparison.
