# VBA-challenge
UCI Data Analysis Bootcamp: Challenge 2 with VBA

Data includes historical data on stocks for the years 2018 - 2020.
Data Columns:
- Ticker
- Date
- Opening Price
- High Price
- Low Price
- Closing Price
- Volume

Utilized the For loops and conditional statements to create an aggregated data table for each unique ticker.
Aggregated values include:
- Yearly Change: stock's closing price at the end of the year - stock's opening price at the beginning of the year
- Percent Yearly Change: (stock's closing price at the end of the year - stock's opening price at the beginning of the year)/opening price
- Total Volume: sum of volumes

Added color formatting to our aggregated table:
- Green: increase in stock price
- Red: decrease in stock price

An additional for loop and conditional statements were used to find the stock tickers with the greatest percent increase and decrease in price as well as the stock with the largest total volume.
