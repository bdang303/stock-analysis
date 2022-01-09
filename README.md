# Stock Analysis with VBA

## Overview of Project
A dataset of stock tickers and prices across a period of two years was the focus of this analysis. Data points such as opening price, closing price, and volume for each day of the year was available as part of the data and the intent was to be able to analyze specific tickers as well as all others to determine returns.

### Purpose
The user was interested in investing in a specific company “DQ” and looking to determine whether or not the data from 2017 and 2018 reflected reasonable returns. In addition, they want to analyze a number of other stocks to see if their returns were also worth a look. Rather than scour through the thousands of rows of data to calculate this, we are able to utilize VBA to write a script to come up with outputs for this

## Results

### Results of Stocks from 2017 vs. 2018
Following the analysis, the specific ticker “DQ” that user was interested in had highest returns in 2017 (199.4%) but however had the worst return rates of all stocks analyzed in 2018. If there was another ticker that we’d recommend investing in it would be ticker “ENPH” which had positive returns in 2017 (129.5%) and 2018 (81.9%)

### Results of 2017 Stock Prices
![2017 Stock Returns](https://github.com/bdang303/stock-analysis/blob/main/AllStocks2017.png)

### Results of 2018 Stock Prices
![2018 Stock Returns](https://github.com/bdang303/stock-analysis/blob/main/AllStocks2018.png)

### VBA Script & Calculations
Using multiple “For” loops, we were able to have the script run through the data set for each of the worksheets that contained data for each of the years. A “For” loop “For i = 0 To 11” was used to run through and find each of the tickers in the data set, and another for loop “For i = 2 To RowCount” was used go through all the rows of data. To ensure we were pulling the current & not new ticket in the data, “IF” statements with 2 critiera were used. 

#### To Determine Starting Prices

          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
          
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

#### To Determine Ending Prices

          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

### Refactoring of Code
By refactoring the initial script, we were able to reduce the run time from 0.304 seconds:
![Initial Code Run Time](https://github.com/bdang303/stock-analysis/blob/main/IntiialScriptRunTime.png)

To 0.074 seconds:

![Refactored Code Run Time](https://github.com/bdang303/stock-analysis/blob/main/RefactoredRunTime.png)


## Summary
### What are the advantages or disadvantages of refactoring code? By refactoring the initial code, statements and logic were significantly less than what was in the original code which could help others better understand the logic when viewing and trying to understand each statement. A disadvantage of going through the process of refactoring is ensuring that the code does the same actions as the initial code that was already working. 
### How do these pros and cons apply to refactoring the original VBA script? One of the main advantages were that we’re able to reduce the overall run time of the script and not require as much memory to execute
![image](https://user-images.githubusercontent.com/93288351/148703970-4cea532c-2e98-4925-aec9-d2a1c4b7d7e5.png)
