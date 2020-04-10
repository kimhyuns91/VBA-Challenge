# VBA-Challenge
Stock Analysis Using VBA
  
# Goal

Given the daily data of the multiple stock data per year, analyze each stock to find the stock to invest and avoid


# Assumptions made
 
  1. Data is sorted out by date starting from lowest to highest for each stock
  
  2. Format of the data is consistent
  
  3. Stocks that had "0" as open value in the beginning of the year, but had a closing value at the end of the year was considered as stock      that had been created in middle of the year and the first opening value that had greater than "0" was used as opening value
  
  
# Method of Analysis

  Only one set of code is written out, but results are shown in each of the tab.
   
  Initially, a new table is created to summarize the yearly information for each stock. For loop and if statement is used to iterate and capture the first opening and last opening value. Afterwards, the yearly changed and percent changed are calculated and presented. Simultaneously, total stock volume for each stock is summed for the year. 
  
  With the summary table showing the information such as change of value and number of stocks, I am able to grab the stock that had biggest change and number of volume. 
  
# Conclusion  

  Looking at the max/min % changes, I can say which stock was most profitable an least porfitable if the stock was purchased from the start until the end of the year. Along with such conclusion, one of the most popular stock can be shown with the stock with the greatest total volume. Maximum % change not matching greatest total number shows that majority of people do not purchase the most profitable stock for the year and it is not easy to predict the stock market.
  
  Additional notes to understand is that these results do not tel the full story. Since the analysis is from beginning to end, it is not the best method for people who are investing in stocks for short terms. Stocks that has the greatest change for the year will not necessarily have greatest change if the duration of the change was different. 
  

  
